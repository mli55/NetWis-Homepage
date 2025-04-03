import os
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import zipfile

def download_jpg_images(url, output_folder="images"):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # 设置请求头，模仿浏览器
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/95.0.4638.69 Safari/537.36"
    }
    
    response = requests.get(url, headers=headers)
    response.raise_for_status()  # 检查请求是否成功
    soup = BeautifulSoup(response.text, "html.parser")
    
    image_tags = soup.find_all("img")
    downloaded_images = []
    
    for img in image_tags:
        img_url = img.get("src")
        if not img_url:
            continue
        
        # 转换为绝对路径
        img_url = urljoin(url, img_url)
        
        # 只处理 jpg 或 jpeg 图片
        if not (img_url.lower().endswith(".jpg") or img_url.lower().endswith(".jpeg")):
            continue
        
        img_name = os.path.basename(img_url.split("?")[0])  # 去除 URL 参数
        img_path = os.path.join(output_folder, img_name)
        try:
            img_data = requests.get(img_url, headers=headers).content
            with open(img_path, "wb") as f:
                f.write(img_data)
            print(f"下载成功: {img_url}")
            downloaded_images.append(img_path)
        except Exception as e:
            print(f"下载失败: {img_url} 错误信息: {e}")
    
    return downloaded_images

def create_zip(folder, zip_name="images.zip"):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for foldername, subfolders, filenames in os.walk(folder):
            for filename in filenames:
                file_path = os.path.join(foldername, filename)
                zipf.write(file_path, os.path.relpath(file_path, folder))
    print(f"已创建 ZIP 文件: {zip_name}")

if __name__ == "__main__":
    website_url = "https://sites.google.com/ncsu.edu/netwis/people"
    images_folder = "downloaded_images"
    download_jpg_images(website_url, images_folder)
    # create_zip(images_folder)