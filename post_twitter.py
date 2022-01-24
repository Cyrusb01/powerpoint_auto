"""
This script imports the generate_images file
Then posts the three latest slides to the twitter
"""


import tweepy
from requests_oauthlib import OAuth1Session
import os
import json
import generate_images

API_KEY = os.getenv("BFC_TWIT_API_KEY")
API_SECRET_KEY = os.getenv("BFC_TWIT_API_SECRET_KEY")

ACCESS_TOKEN = os.getenv("BFC_TWIT_ACCESS_TOKEN")
ACCESS_TOKEN_SECRET = os.getenv("BFC_TWIT_ACCESS_TOKEN_SECRET")


oauth = OAuth1Session(
    API_KEY,
    client_secret=API_SECRET_KEY,
    resource_owner_key=ACCESS_TOKEN,
    resource_owner_secret=ACCESS_TOKEN_SECRET,
)



url_media = "https://upload.twitter.com/1.1/media/upload.json"
url_text = "https://api.twitter.com/1.1/statuses/update.json"





def get_media_id(image_path):
    
    files = {"media" : open(image_path, 'rb')}
    req_media = oauth.post(url_media, files = files)

    if req_media.status_code != 200:
        print ("image app fail: %s", req_media.text)
        exit()

    media_id = json.loads(req_media.text)['media_id']
    print ("Media ID: %d" % media_id)

    return media_id


media_id_1 = get_media_id("slide_images/slide_1.png")
media_id_2 = get_media_id("slide_images/slide_2.png")
media_id_3 = get_media_id("slide_images/slide_3.png")

ids = str(media_id_1) + "," + str(media_id_2) + "," + str(media_id_3)

message = "Blockforce MultiStrategy Hedge Fund Performance and Portfolio Update. Please sign up for our monthly market commentary here https://blockforcecapital.com/"

params = {'status': message,  "media_ids": ids}
req_text = oauth.post(url_text, params = params)

if req_text.status_code != 200:
    print ("text app fail: %s", req_text.text)
    exit()

print ("Posted to Twitter")
