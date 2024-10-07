import win32com.client
import configparser
import pythoncom
from datetime import datetime
import time

client = win32com.client.Dispatch("XA_Session.XASession")
client.ConnectServer("demo.ebstsec.co.kr", 20001)
# 모의투자 서버 주소 : 'demo.ebestsec.co.kr'
# 포트 번호 : '20001'
print(client.ConnectServer("demo.ebstsec.co.kr", 20001))

class XASession:
    # 로그인 상태를 확인하기 위한 클래스 변수
    logic_state = 0

    def onLogin(self, code, msg):
        """
        로그인 시도 후 호출되는 이벤트.
        code가 000이면 로그인 성공
        """
        if code == '000':
            print(code, msg)
            XASession.login_state = 1
        else:
            print(code, msg)

    def onDisconnect(self):
        """
        서버와 연결이 끊어지면 발생하는 이벤트.
        """
        print("Session Disconnected")
        XASession.login_state = 0

class EBest:

    QUERY_LIMIT_10MIN = 200
    LIMIT_SECONDS = 600 #10min



    def __init__(self, mode=None):
        """
        config.ini 파일을 로드해 사용자, 서버 정보 저장
        query_cnt : 10분당 200개의 TR 수행을 관리하기 위한 리스트
        xa_session_client : XASession 객체
        :param mode:str - 모의 서버는 DEMO 실서버는 PROD로 구분
        """
        if mode not in ["PROD", "DEMO"]:
            raise Exception("Need to run_mode(PROD or DEMO)")
        run_mode = "EBEST_" + mode
        config = configparser.ConfigParser()
        config.read('conf/config.ini')
        self.user = config[run_mode]['user']
        self.passwd = config[run_mode]['password']
        self.cert_passwd = config[run_mode]['cert_passwd']
        self.host = config[run_mode]['host']
        self.port = config[run_mode]['port']
        self.account = config[run_mode]['account']

        self.xa_session_client = win32com.client.DispatchWithEvents("XA_Session.XASession", XASession)
        self.query_cnt = []
        
        def login(self):
            self.xa_session_client.ConnectServer(self.host, self.port)
            self.xa_session_client.Login(self.user, self.passwd, self.cert_passwd, 0, 0)
            while XASession.login_state == 0:
                pythoncom.PumpWaitingMessages()

        def logout(self):
            #result = self.xa_session_client.Logout()
            #if result:
            XASession.login_state = 0
            self.xa_session_client.DisconnectServer()

        def _execute_query(self, res, in_block_name, out_block_name, *out_fields, **set_fields):
            """
            TR 코드를 실행하기 위한 메소드입니다.
            :param res:str 리소스 이름(TR)
            :param in_block_name:str 인 블록 이름
            :param out_block_name:str 아웃 블록 이름
            :param out_params:list 출력 필드 리스트
            :param in_params:dict 인 블록에 설정할 필드 딕셔너리
            :param result:list 결과를 list에 담아 반환
            """
            time.sleep(1)
            print("current query cnt:", len(self.query_cnt))
            print(res, in_block_name, out_block_name)
            while len(self.query_cnt) >= EBest.QUERY_LIMIT_10MIN:
                time.sleep(1)
                print("waiting for execute query... current query cnt:", len(self.query_cnt))
                self.query_cnt = list(filter(lambda x: (datetime.today() - x).total_seconds()
                                             < EBest.LIMIT_SECONDS, self.query_cnt))
            
            xa_query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQuery)
            xa_query.LoadFromResFile(XAQuery.RES_PATH + res+".res")
            #in_block_name 세팅
            for key, value in set_fields.items():
                xa_query.SetFieldData(in_block_name, key, 0, value)
            errorcode = xa_query_Request(0)
            
            # 요청 후 대기
            waiting_cnt = 0
            while xa_query.tr_run_state == 0:
                waiting_cnt += 1
                if waiting_cnt % 100000 == 0:
                    print("Waiting...", self.xa_session_client.GetLastError())
                    pythoncom.PumpWaitingMessages()
            
            result = []
            count = xa_query.GetBlockCount(out_block_name)

            for i in range(count):
                item = {}
                for field in out_fields:
                    value = xa_query.GetFieldData(out_block_name, field, i)
                    item[field] = value
                result.append(item)
            
            # 제약시간 체크
            XAQuery.tr_run_state = 0
            self.query_cnt.append(datetime.today())

            # 영문 필드명을 한글 필드명으로 변환
            for item in result:
                for field in list(item.keys()):
                    if getattr(Field, res, None):
                        res_field = getattr(Field, res, None)
                        if out_block_name in res_field:
                            field_hname = res_field[out_block_name]
                            if field in field_hname:
                                item[field_hname[field]] = item[field]
                                item.pop(field)
            
            return result




            

class XAQuery:
    RES_PATH = "C:\\eBEST\\xingAPI\\Res\\"
    tr_run_state = 0

    def OnReceiveData(self, code):
        print("OnReceiveData", code)
        XAQuery.tr_run_state = 1

    def OnReceiveMessage(self, error, code, message):
        print("OnreceiveMessage", error, code, message)