import requests


class Post():
    def __init__(self): # 인잇 생성자와 셀프 무조건 들어가야함
        return
    
    def post_message(self, channel, text):
        myToken = "xoxb-3162802854482-3643669615106-zrrUHSu6qTmcQ361pLUvPvou"
        response = requests.post("https://slack.com/api/chat.postMessage",
            headers={"Authorization": "Bearer "+myToken},
            data={"channel": channel ,"text": text})
        
 
    