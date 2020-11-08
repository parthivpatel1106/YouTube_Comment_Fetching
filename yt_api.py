import os
from googleapiclient.discovery import build
from ytapi_key import api_key
import xlsxwriter
youtube=build('youtube','v3',developerKey=api_key)
print(type(youtube))
print('enter channel name:')
channel_name= input()
req_channel=youtube.search().list(q=channel_name, part='snippet', type='channel', order='relevance', maxResults=1).execute()
#res_channel=req_channel.execute()
channel_id=req_channel['items'][0]['id']['channelId']
req_channel2=youtube.channels().list(part='contentDetails', id=channel_id, maxResults=1).execute()
#res_channel2=req_channel2.execute()
playlist_id=req_channel2['items'][0]['contentDetails']['relatedPlaylists']['uploads']
print("enter the total number of videos")
totalVideos=int(input())
req_videos=youtube.playlistItems().list(playlistId=playlist_id, part='snippet', maxResults=totalVideos).execute()
#res_videos=req_videos.execute()
print("Enter video number")
videos=int(input())
video_name=req_videos['items'][videos]['snippet']['title']
video_id=req_videos['items'][videos]['snippet']['resourceId']['videoId']
print(video_name,"\n", video_id)
comments=[]
nextPage_token=None
while 1:
    req_comment=youtube.commentThreads().list(part='snippet',videoId=video_id, maxResults=100, order='relevance', pageToken=nextPage_token).execute()
    nextPage_token=req_comment.get('nextPageToken')
    for item in req_comment['items']:
        comments.append(item['snippet']['topLevelComment']['snippet']['textOriginal'])
    print(comments)
    print(len(comments))
    if nextPage_token is None:
        break 
workbook=xlsxwriter.Workbook(video_id + " " + channel_name + '.xlsx')
worksheet=workbook.add_worksheet()
row=0
column=0
for comment in comments:
    worksheet.write(row,column,comment)
    row+=1
workbook.close()