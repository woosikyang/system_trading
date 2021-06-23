from slacker import Slacker
from configs import *
import requests


def post_message(token, channel, text):
    response = requests.post("https://slack.com/api/chat.postMessage",
                             headers={"Authorization": "Bearer " + token},
                             data={"channel": channel, "text": text}
                             )
    print(response)


text = 'Hello World'
post_message(slack_api_token, "#주식", text)