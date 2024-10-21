import time
import os
import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys

# Initialize the Selenium WebDriver to extract data
driver_path = 'C:\\Users\\amola\\OneDrive\\Documents\\pankaj\\Ds\\applied jobs\\GoQuest Media\\chromedriver.exe' # Your Chrome Driver Path
service = Service(driver_path)
driver = webdriver.Chrome(service=service)

# Extract all video from the channel handle, we observe that the channel handle has multiple playlist to extract video from playlist
# we need to gather the all Playlist urls

def playlist_videos_urls(channel_url):
    # To Open the channel playlist page we provided path
    driver.get(channel_url + '/playlists')
    time.sleep(5) 

    # Fetch all Playlist Url from the channel handle
    Playlist_elements = driver.find_elements(By.XPATH, '//ytd-grid-playlist-renderer//yt-formatted-string/a')  
    print("Scraped the Playlist Url for the the channel hanndle")
    Playlist_urls = [element.get_attribute('href') for element in Playlist_elements]
    
    return Playlist_urls


# Collect all urls on the youtube channel handle
def fetch_all_Videos_urls(Playlist_urls):
    # Open the channel playlist page
    Overall_video_urls=[]
    for playlist in Playlist_urls:
        driver.get(playlist) # open the Playlist videos
        time.sleep(5)
        video_elements = driver.find_elements(By.XPATH, '//ytd-playlist-video-renderer//ytd-thumbnail/a')
        print("Scraped the all video Urls from the the Playlists")
        video_urls = [element.get_attribute('href') for element in video_elements]
        Overall_video_urls.extend(video_urls)
    
    return Overall_video_urls


def extract_video_data(driver, video_url):
    driver.get(video_url)
    time.sleep(5)

    # Fetch video details
    video_id = video_url.split('list=')[-1]

    try:
        driver.execute_script(f"window.scrollTo(0, 500);")
        time.sleep(2)
    
        comment_count_element = driver.find_element(By.XPATH, '//*[@id="count"]/yt-formatted-string/span[1]')
        comment_count = comment_count_element.text
    except Exception as e:
        comment_count = "Comments Disabled"
        
    try:
        show_more_button = driver.find_element(By.XPATH, '//*[@id="description-inline-expander"]/tp-yt-paper-button')
        show_more_button.click()
        time.sleep(2)
    except Exception as a:
        print(f"No 'Show more' button found. The description is already expanded.{a}") 

    try:
        title = driver.find_element(By.XPATH, '//ytd-watch-metadata//yt-formatted-string').text
    except:
        title=None
    try:
        description = driver.find_element(By.XPATH, '//*[@id="description-inline-expander"]').text 
    except:
        description=None
    try:
        published_date = driver.find_element(By.XPATH, '//*[@id="info"]/span[3]').text
    except:
        published_date=None
    try:
        view_count = driver.find_element(By.XPATH, '//*[@id="info"]/span[1]').text
    except:
        view_count=None
    try:   
        duration = driver.find_element(
        By.XPATH, '//*[@id="overlays"]/ytd-thumbnail-overlay-time-status-renderer/div[1]/badge-shape').text
    except:
        duration=None
    try:
        thumbnail_url = driver.find_element(By.XPATH, '//*[@id="thumbnail"]/yt-image/img').get_attribute('src')
    except:
        thumbnail_url=None
    try:
        like_count = driver.find_element(
        By.XPATH, '//*[@id="top-level-buttons-computed"]/segmented-like-dislike-button-view-model/yt-smartimation/div/div/like-button-view-model/toggle-button-view-model/button-view-model/button/div[2]').text
    except:
        like_count=None

    print(f"Extracting Video data from the title:{title} ")
    
    return {
        'video_id': video_id,
        'title': title,
        'description': description,
        'published_date': published_date,
        'view_count': view_count,
        'like_count': like_count,
        'comment_count': comment_count,
        'duration': duration,
        'thumbnail_url': thumbnail_url,
        'video_url': video_url
    }



def extract_comments(driver, video_url):
    fetch_comment_data=[]
    video_id = video_url.split('list=')[-1]
    driver.get(video_url)
    time.sleep(5)  # Wait for the page to load

    driver.execute_script(f"window.scrollTo(0, 500);")
    time.sleep(2)
    total_comments_fetched = 0

    try:
        sort_button  = driver.find_element(By.XPATH, '//yt-sort-filter-sub-menu-renderer//tp-yt-paper-button')
        sort_button.click()
        time.sleep(1)
    
        newest_first = driver.find_element(
            By.XPATH, '/html/body/ytd-app/div[1]/ytd-page-manager/ytd-watch-flexy/div[5]/div[1]/div/div[2]/ytd-comments/ytd-item-section-renderer/div[1]/ytd-comments-header-renderer/div[1]/div[2]/span/yt-sort-filter-sub-menu-renderer/yt-dropdown-menu/tp-yt-paper-menu-button/tp-yt-iron-dropdown/div/div/tp-yt-paper-listbox/a[2]/tp-yt-paper-item')
        newest_first.click()
        time.sleep(1)
        
        SCROLL_PAUSE_TIME = 2  # Time to pause when scrolling
        delay = 2
        scrolling = True
        scrolling_attempt = 2
        fetched_comments_set = set()
        # Scroll down to load comments
        last_height = driver.execute_script("return document.documentElement.scrollHeight")

        while scrolling == True and total_comments_fetched < 100:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
            time.sleep(SCROLL_PAUSE_TIME)

            try:
                # comment_id = WebDriverWait(driver, delay).until(
                #         EC.presence_of_all_elements_located((By.XPATH, 'id')))
                all_comment_elements = WebDriverWait(driver, delay).until(
                    EC.presence_of_all_elements_located((By.XPATH, '//*[@id="comment"]'))
                )
                        
                all_usernames = WebDriverWait(driver, delay).until(
                        EC.presence_of_all_elements_located((By.XPATH, '//*[@id="author-text"]')))
                all_comments = WebDriverWait(driver, delay).until(
                            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="content-text"]')))
                all_published_dates = WebDriverWait(driver, delay).until(
                            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="published-time-text"]/a')))
                all_like_count = WebDriverWait(driver, delay).until(
                            EC.presence_of_all_elements_located((By.XPATH, '//*[@id="vote-count-middle"]')))
                try:
                    # we'll try to get only last 20 elements, because youtube loads 20 comments per scroll
                    new_comment_elements = all_comment_elements[:100]
                    new_comments = all_comments[:100]
                    new_usernames = all_usernames[:100]
                    new_published_dates = all_published_dates[:100]
                    new_like_count = all_like_count[:100]
                except:
                    print("could not get last 20 elements")

                for (comment_element,username, comment,published_date,like_count) in zip(
                    new_comment_elements,new_usernames, new_comments,new_published_dates,new_like_count):
                    comment_text = comment.text
                    comment_id = comment_element.get_attribute('data-id') or f"comment-{total_comments_fetched}"
                    if comment_text not in fetched_comments_set:
                        current_comment = {"video_id":video_id,
                                            "comment_id":comment_id,
                                            "author_name": username.text,
                                            "published_date":published_date.text,
                                            "comment_text": comment.text,
                                            "like_count":like_count.text,
                                            "replies": []}
                        print(f"video_id:{video_id}\nauthor_name : {username.text}\ncomment_text : {comment.text}")
                        fetch_comment_data.append(current_comment)  # here we'll store comments
                        fetched_comments_set.add(comment_text)
                        total_comments_fetched += 1

                        if total_comments_fetched >= 100:
                            scrolling = False
                            break

                driver.find_element(By.TAG_NAME, "body").send_keys(Keys.END)
                time.sleep(SCROLL_PAUSE_TIME)

                new_height = driver.execute_script("return document.documentElement.scrollHeight") # calculate current position
                time.sleep(SCROLL_PAUSE_TIME)

                if new_height == last_height: 
                    scrolling_attempt -= 1
                    print(f"scrolling attempt {scrolling_attempt}")
                    if(scrolling_attempt == 0):
                        scrolling = False # this will break while loop
                last_height = new_height #
                
            except Exception as a:
                print(f"error while trying to load comments:{a}")
                scrolling_attempt -= 1
                if scrolling_attempt == 0:
                    scrolling = False
    except Exception as c:
        print("comment section disabled")    
        current_comment = {"video_id":video_id,
                            "comment_id":"comment disabled",
                            "author_name": "comment disabled",
                            "published_date": "comment disabled",
                            "comment_text": "comment disabled",
                            "like_count":"comment section disabled",
                            "replies":"disabled"}
        fetch_comment_data.append(current_comment)
    return fetch_comment_data

def fetch_video_data(Overall_video_urls):
    video_data=[]
    for video_url in Overall_video_urls:
        video_info=extract_video_data(driver, video_url)
        video_data.append(video_info)
    return video_data

def fetch_comment_data(Overall_video_urls):
    comment_data=[]
    for video_url in Overall_video_urls:
        comment_info=extract_comments(driver, video_url)
        comment_data.extend(comment_info)
    return comment_data




def main(channel_url):
    # Step 1: Fetch video data
    Playlist_urls=playlist_videos_urls(channel_url)
    print(f"Total Playlist in the channel Handle {channel_url} is : {len(Playlist_urls)}")
    Overall_video_urls=fetch_all_Videos_urls(Playlist_urls)
    print(f"Total videos in the channel Handle {channel_url} is : {len(Overall_video_urls)}")
    videos_data = fetch_video_data(Overall_video_urls)
    video_df = pd.DataFrame(videos_data)
    print(f"Total videos data in the channel Handle {channel_url} is : {len(video_df)}")
    
    # Save video data to the Excel file first
    with pd.ExcelWriter('YouTube_Data_file.xlsx', mode='w', engine='openpyxl') as writer:
        video_df.to_excel(writer, sheet_name='Video Data', index=False)
        print("Video data saved successfully.")

    # Step 2: Fetch comments data
    comment_data = fetch_comment_data(Overall_video_urls)
    comments_df = pd.DataFrame(comment_data)
    print(f"Total comment data in the channel Handle {channel_url} is : {len(comments_df)}")
    
    with pd.ExcelWriter('YouTube_Data_file.xlsx', mode='a', engine='openpyxl') as writer:
        comments_df.to_excel(writer, sheet_name='Comments Data', index=False)
        print("Comments data saved successfully.")
    # # Step 3: Save data to Excel
    # with pd.ExcelWriter('YouTube_Data_file.xlsx') as writer:
    #     video_df.to_excel(writer, sheet_name='Video Data', index=False)
    #     comments_df.to_excel(writer, sheet_name='Comments Data', index=False)
    
    print("Data saved to YouTube_Data.xlsx")

if __name__ == "__main__":
    # Replace with the actual YouTube channel URL
    channel_url = 'https://www.youtube.com/@channelhandle/playlists'
    main(channel_url)

    # Close the Selenium WebDriver
    
