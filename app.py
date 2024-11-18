import os
from flask import Flask, jsonify, request, send_file, send_from_directory
from googleapiclient.discovery import build
import pandas as pd
from io import BytesIO

app = Flask(__name__, static_folder='web')

# YouTube API 키 설정
API_KEY = os.getenv('API_KEY')  # Render 환경 변수에서 API_KEY 가져오기
youtube = build('youtube', 'v3', developerKey=API_KEY)

@app.route('/')
def index():
    return send_from_directory(app.static_folder, 'index.html')

@app.route('/<path:path>')
def serve_file(path):
    return send_from_directory(app.static_folder, path)

@app.route('/fetch-data', methods=['POST'])
def fetch_data():
    """채널의 동영상 데이터 가져오기"""
    data = request.json
    channel_id = data.get('channelId')
    start_date = data.get('startDate')
    end_date = data.get('endDate')

    if not channel_id or not start_date or not end_date:
        return jsonify({"error": "Missing required parameters"}), 400

    try:
        all_videos = []
        next_page_token = None

        # 모든 동영상 가져오기
        while True:
            response = youtube.search().list(
                part='id,snippet',
                channelId=channel_id,
                publishedAfter=f"{start_date}T00:00:00Z",
                publishedBefore=f"{end_date}T23:59:59Z",
                maxResults=50,
                pageToken=next_page_token
            ).execute()

            for item in response['items']:
                video_id = item['id'].get('videoId')
                title = item['snippet']['title']
                published_at = item['snippet']['publishedAt']

                if video_id:
                    # 동영상의 통계 데이터 가져오기
                    stats_response = youtube.videos().list(
                        part='statistics',
                        id=video_id
                    ).execute()

                    stats = stats_response['items'][0]['statistics']
                    all_videos.append({
                        "videoId": video_id,
                        "title": title,
                        "publishedAt": published_at,
                        "viewCount": stats.get('viewCount', 0),
                        "likeCount": stats.get('likeCount', 0),
                        "commentCount": stats.get('commentCount', 0)
                    })

            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break

        return jsonify({"videos": all_videos})

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/fetch-comments', methods=['POST'])
def fetch_comments():
    """특정 동영상의 댓글 가져오기"""
    data = request.json
    video_id = data.get('videoId')

    if not video_id:
        return jsonify({"error": "Missing videoId parameter"}), 400

    try:
        all_comments = []
        next_page_token = None

        # 모든 댓글 가져오기
        while True:
            response = youtube.commentThreads().list(
                part='snippet',
                videoId=video_id,
                maxResults=50,
                pageToken=next_page_token
            ).execute()

            for item in response['items']:
                top_comment = item['snippet']['topLevelComment']['snippet']
                all_comments.append({
                    "videoId": video_id,
                    "author": top_comment['authorDisplayName'],
                    "comment": top_comment['textDisplay'],
                    "publishedAt": top_comment['publishedAt']
                })

            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break

        return jsonify(all_comments)

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/download-excel', methods=['POST'])
def download_excel():
    """동영상 및 댓글 데이터를 Excel로 다운로드"""
    data = request.json
    videos = data.get('videos', [])

    all_comments = []

    # 모든 동영상의 댓글 가져오기
    for video in videos:
        video_id = video['videoId']
        next_page_token = None

        while True:
            response = youtube.commentThreads().list(
                part='snippet',
                videoId=video_id,
                maxResults=50,
                pageToken=next_page_token
            ).execute()

            for item in response['items']:
                top_comment = item['snippet']['topLevelComment']['snippet']
                all_comments.append({
                    "videoId": video_id,
                    "videoTitle": video['title'],
                    "author": top_comment['authorDisplayName'],
                    "comment": top_comment['textDisplay'],
                    "publishedAt": top_comment['publishedAt']
                })

            next_page_token = response.get('nextPageToken')
            if not next_page_token:
                break

    # 데이터프레임 생성
    videos_df = pd.DataFrame(videos)
    comments_df = pd.DataFrame(all_comments)

    # 엑셀 파일 생성
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        videos_df.to_excel(writer, index=False, sheet_name='Videos')
        comments_df.to_excel(writer, index=False, sheet_name='Comments')
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name='youtube_data.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == '__main__':
    # Render가 제공하는 PORT 환경 변수 사용
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
