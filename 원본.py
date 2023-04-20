Project_Name = """Schedule Project
배포 버젼 9"""

import os.path, sys, io, json, requests

import tkinter as  tk
import tkinter.font
from tkinter import filedialog

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

class Kakaotalk:
    "나에게 카카오톡 메세지 보내기"
    def __init__(self, sheet_data, api):
        self.__is_kakao = sheet_data["iskakaotalk"]
        self.token = json.loads(sheet_data["kakao_token"])
        self.__REST_API_KEY = api["kakao REST API"]

    def Send(self, text):
        "텍스트 = 나에게 카톡 보낼 내용"
        url = "https://kapi.kakao.com/v2/api/talk/memo/default/send"

        headers = {
            "Authorization" : "Bearer {}".format(self.token["access_token"])
        }

        data = {
            "template_object" : json.dumps({
                "object_type" : "text",
                "text" : text,
                "link" : {
                    "web_url" : "https://www.naver.com",
                    "mobile_web_url" : "https://m.naver.com"
                }
            })
        }

        response = requests.post(url, headers=headers, data=data)
        return response.status_code

    def Token_Reissue(self):
        url = "https://kauth.kakao.com/oauth/token"

        data = {
            "grant_type" : "refresh_token",
            "client_id" : self.__REST_API_KEY,
            "refresh_token" : self.token["refresh_token"]
        }

        response = requests.post(url, data=data)
        response = response.json()

        if response.get("access_token"):
            self.token["access_token"] = response["access_token"]
            return self.token
        if not response.get("error"):
            pass
            # 여기다가 새로운 토큰 받기

    def __call__(self, text):
        if not self.__is_kakao:
            return

        result = self.Send(text)
        if result == 200: # 정상적으로 보내짐
            return
        elif result == 401: # 토큰 만료, 토큰 다시 발급
            return_token = self.Token_Reissue()
            self.Send(text)
            return return_token

class GoogleSheet:
    "구글 스프레드시트의 정보를 가져오고 보내는 함수"
    def __init__(self, api):
        self.api = api
        self.__url = api["google_sheet url"]
        self.json_data = None

    def get(self) -> dict:
        '정보를 받아옴 ["token"], ["storage"], ["iskakaotalk"], ["kakao_token"], ["file_id"], ["upload_url"]'
        response = requests.get(self.__url)
        if response.status_code != 200:
            # print("겟 오류 남")
            return

        self.json_data = response.json()
        return self.json_data

    def post(self, value, row=8, column=1):
        data = {
            "row": row,
            "column": column,
            "value": value
        }
        response = requests.post(self.__url, data=data)
        if response.status_code != 200:
            # print("포스트 오류 남")
            return
        # print("보내기 완료")

class Google:
    "구글 드라이드 API 사용"
    def __init__(self, sheet_data, kakao, google_sheet, api):
        # 받아온 정보들
        token_json = json.loads(sheet_data["token"])
        storage_json = json.loads(sheet_data["storage"])
        self.file_id = sheet_data["file_id"]
        self.upload_url = sheet_data["upload_url"]
        self.kakao = kakao
        # 권한 인증 및 토큰 확인
        self.SCOPES = ['https://www.googleapis.com/auth/drive']
        self.FOLDER_ID = api["forder_id"] #위에서 복사한 구글드라이브 폴더의 id
        self.creds = None

        # 이미 발급받은 토큰이 있을 때
        if token_json:
            try:
                # self.creds = Credentials.from_authorized_user_file('token.json', self.SCOPES)
                self.creds = Credentials.from_authorized_user_info(token_json, self.SCOPES)
            except Exception as e:
                print(f"발급받은 토큰에서 문제가 생김: {e}")
                self.creds = None

        # 발급받은 토큰이 없거나 엑세스토큰이 만료되었을 때
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                # flow = InstalledAppFlow.from_client_secrets_file(r'시간표가져오기\storage.json', self.SCOPES)
                flow = InstalledAppFlow.from_client_config(storage_json, self.SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            # 다운받기 뭐하지만 일단은 다운받아두는걸로 하자
            google_sheet.post(self.creds.to_json(), row=2, column=1)

        # 연결 인스턴스 생성
        self.service = build('drive', 'v3', credentials=self.creds)

    def Folder_Make(self, folder_name):
        "드라이브에 폴더를 생성, 여기선 안쓰임"
        file_metadata = {
            'name': folder_name,
            'mimeType': 'application/vnd.google-apps.folder'
        }
        file = self.service.files().create(body=file_metadata, fields='id').execute() # 폴더 실행문
        print('Folder ID: %s' % file.get('id'))
        return file.get('id')
    
    def File_Upload(self):
        "특정 폴더에 파일을 올림, 폴더값은 지금 고정되어있음"
        name = os.path.basename(self.upload_url)

        file_metadata = {'name': name, "parents": [self.FOLDER_ID]}
        media = MediaFileUpload(self.upload_url, resumable=True)
        file = self.service.files().create(body=file_metadata,
                                           media_body=media,
                                           fields='id').execute()
        print('File ID: %s' % file.get('id'))
        self.kakao('File ID: %s' % file.get('id'))
        return file.get('id')

    def File_Download(self, file_path, file_id):
        "다운받을 파일경로, 파일 id"
        file_path = os.path.join(os.path.dirname(sys.argv[0]), "시간표.xlsx")

        request = self.service.files().get_media(fileId=file_id)
        fh = io.FileIO(file_path, "wb")
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print("Download %d%%." % int(status.progress() * 100), flush=True)
        self.kakao("Download %d%%." % int(status.progress() * 100))

    def File_Delete(self, file_id):
        "파일 삭제"
        try:
            self.service.files().delete(fileId=file_id).execute()
            self.kakao(f"File ID:: {file_id} 삭제 완료")
        except Exception:
            pass

def Data_Reading(filePath):
    """
    filePath: 저장된 경로의 txt, json 파일 읽음
    return: str or dict
    """
    with open(filePath, "r", encoding='utf8') as f:
        data = f.read()

    if os.path.splitext(filePath)[-1] == ".json":
        data = json.loads(data)

    return data

def Upload_main(api):
    google_sheet = GoogleSheet(api)
    sheet_data = google_sheet.get()
    kakao = Kakaotalk(sheet_data, api)
    # token = kakao("실험용")
    # if token != None:
    #     google_sheet.post(token, row=6, column=1)
    google_api = Google(sheet_data, kakao, google_sheet, api)

    # 파일id 이용해서 파일 지우기
    google_api.File_Delete(sheet_data["file_id"])

    # 파일 올리고 -> 시트에 파일id 적기
    file_id = google_api.File_Upload()
    google_sheet.post(file_id)

    kakao("업로드 완료")

def Download_main(api):
    google_sheet = GoogleSheet(api)
    sheet_data = google_sheet.get()
    kakao = Kakaotalk(sheet_data, api)
    google_api = Google(sheet_data, kakao, google_sheet, api)

    # 다운로드
    google_api.File_Download("", sheet_data["file_id"])

    kakao("다운로드 완료")

class SettingWindow:
    def __init__(self, win):
        size = (300, 200)
        self.google_sheet = GoogleSheet()
        # 윈도우 창 만들기
        self.win = win
        self.win.title("대충 설정창")
        self.win.geometry(f"{size[0]}x{size[1]}")
        font25 = tk.font.Font(size=25)
        font13 = tk.font.Font(size=13)

        main_label = tk.Label(self.win, text="간단 설정창", font=font25, borderwidth=10)
        main_label.pack()
        serve_label = tk.Label(self.win, text="무엇을 업로드 할건지 정하기", font=font13, borderwidth=10)
        serve_label.pack()

        # 주소 입력창
        self.path_entry = tk.Entry(self.win, width=round(size[0]*0.12))
        self.path_entry.bind("<Button-1>", self.entry_Func)
        self.path_entry.pack()

        # 버튼
        self.button = tk.Button(self.win, text="저장하고 닫기", command=self.button_Func)
        self.button.pack()

        self.win.mainloop()

    def entry_Func(self, event):
        entry_dirName = filedialog.askopenfilename()
        if entry_dirName:
            self.path_entry.delete(0, len(self.path_entry.get()))
            self.path_entry.insert(0, entry_dirName)
        # self.save() 나중에 저장 따로 대체
    
    def button_Func(self):
        self.google_sheet.post(self.path_entry.get(), row=10, column=1)
        self.win.destroy()

if __name__ == "__main__":
    name = os.path.basename(sys.argv[0])
    # dataAPI.json = {
        # "kakao REST API": "..."
        # "google_sheet url": "..."
        # "forder_id": "..."
    # }
    api = Data_Reading("dataAPI.json")
    if ("setting".lower() in name.lower()) or ("설정" in name):
        win = tk.Tk()
        SettingWindow(win)
    elif ("upload".lower() in name.lower()) or ("업로드" in name):
        Upload_main(api)
    elif ("download".lower() in name.lower()) or ("다운" in name):
        Download_main(api)