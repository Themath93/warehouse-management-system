
from PyKakao import Message



## 카카오톡 메시지
api = Message(service_key = "fb7a4a68eab473037f341e7cd7c73973")
auth_url = api.get_url_for_generating_code()
print(auth_url)
url = "https://kauth.kakao.com/oauth/authorize?scope=talk_message&response_type=code&redirect_uri=https%3A%2F%2Flocalhost%3A5000&through_account=true&client_id=fb7a4a68eab473037f341e7cd7c73973&app_type=web"
access_token = api.get_access_token_by_redirected_url(url)
api.set_access_token(access_token)
