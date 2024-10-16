import googleapiclient.discovery
import pandas as pd

# Set up the YouTube API client
def get_youtube_client():
    api_key = "AIzaSyDGEfHflsyRK7pV76BlCbxuaJIVzgA4iF0"  # Replace with your valid API key
    return googleapiclient.discovery.build("youtube", "v3", developerKey=api_key)

# Get the channel ID using the handle
def get_channel_id_by_handle(youtube, handle):
    request = youtube.search().list(
        part="snippet",
        q=handle,  # Search using the handle (e.g., "@channelhandle")
        type="channel",
        maxResults=1  # Fetch the first result
    )
    response = request.execute()

    # Check if any items were returned
    if 'items' in response and len(response['items']) > 0:
        return response['items'][0]['snippet']['channelId']  # Extract the channel ID
    else:
        raise ValueError(f"No channel found for handle: {handle}")

# Fetch detailed video data, including view count, like count, etc.
def get_videos_from_channel(youtube, channel_id):
    videos = []
    next_page_token = None

    while True:
        request = youtube.search().list(
            part="snippet",
            channelId=channel_id,
            maxResults=50,  # Fetch 50 videos at a time
            order="date",  # Order by most recent
            pageToken=next_page_token
        )
        response = request.execute()

        video_ids = [item["id"]["videoId"] for item in response['items'] if item["id"]["kind"] == "youtube#video"]

        # Fetch additional video details like view count, duration, etc.
        details_request = youtube.videos().list(
            part="contentDetails,statistics,snippet",
            id=','.join(video_ids)
        )
        details_response = details_request.execute()

        for item in details_response['items']:
            video_data = {
                "video_id": item["id"],
                "title": item["snippet"]["title"],
                "description": item["snippet"]["description"],
                "publish_date": item["snippet"]["publishedAt"],
                "view_count": item["statistics"].get("viewCount", 0),
                "like_count": item["statistics"].get("likeCount", 0),
                "comment_count": item["statistics"].get("commentCount", 0),
                "duration": item["contentDetails"]["duration"],
                "thumbnail_url": item["snippet"]["thumbnails"]["default"]["url"],
            }
            videos.append(video_data)

        # Handle pagination if there are more results
        next_page_token = response.get("nextPageToken")
        if not next_page_token:
            break

    return videos

# Fetch comments and replies for each video
def get_comments_for_video(youtube, video_id):
    comments = []
    next_page_token = None

    while True:
        request = youtube.commentThreads().list(
            part="snippet,replies",
            videoId=video_id,
            maxResults=100,  # Limit to the latest 100 comments
            pageToken=next_page_token
        )
        response = request.execute()

        for item in response["items"]:
            top_comment = item["snippet"]["topLevelComment"]["snippet"]
            comment_data = {
                "video_id": video_id,
                "comment_id": item["snippet"]["topLevelComment"]["id"],
                "comment_text": top_comment["textOriginal"],
                "author": top_comment["authorDisplayName"],
                "published_at": top_comment["publishedAt"],
                "like_count": top_comment.get("likeCount", 0),
                "reply_to": None  # Top-level comments have no parent
            }
            comments.append(comment_data)

            # If there are replies, include them
            if item.get("replies"):
                for reply in item["replies"]["comments"]:
                    reply_data = {
                        "video_id": video_id,
                        "comment_id": reply["id"],
                        "comment_text": reply["snippet"]["textOriginal"],
                        "author": reply["snippet"]["authorDisplayName"],
                        "published_at": reply["snippet"]["publishedAt"],
                        "like_count": reply["snippet"].get("likeCount", 0),
                        "reply_to": item["snippet"]["topLevelComment"]["id"]  # Reply to the top-level comment
                    }
                    comments.append(reply_data)

        next_page_token = response.get("nextPageToken")
        if not next_page_token:
            break

    return comments

# Save video and comment data to Excel
def save_to_excel(video_data, comment_data, output_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Convert video data to a DataFrame and write to Sheet 1
        video_df = pd.DataFrame(video_data)
        video_df.to_excel(writer, sheet_name='Video Data', index=False)

        # Convert comment data to a DataFrame and write to Sheet 2
        comment_df = pd.DataFrame(comment_data)
        comment_df.to_excel(writer, sheet_name='Comments Data', index=False)

# Main function
def main():
    youtube = get_youtube_client()
    channel_handle = "https://www.youtube.com/@YahooBaba"  # Replace with the actual channel handle

    try:
        # Fetch the channel ID using the handle
        channel_id = get_channel_id_by_handle(youtube, channel_handle)
        print(f"Channel ID: {channel_id}")

        # Fetch video data
        videos = get_videos_from_channel(youtube, channel_id)

        # Fetch comments for all videos
        all_comments = []
        for video in videos:
            video_id = video["video_id"]
            comments = get_comments_for_video(youtube, video_id)
            all_comments.extend(comments)  # Add comments to the overall list

        # Save the data to an Excel file
        save_to_excel(videos, all_comments, 'youtube_channel_data.xlsx')

        print("Data has been saved to youtube_channel_data.xlsx")

    except ValueError as e:
        print(e)

if __name__ == "__main__":
    main()
