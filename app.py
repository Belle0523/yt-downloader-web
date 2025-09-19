import os
import re
import json
import pandas as pd
from pathlib import Path
from flask import Flask, render_template, request, send_from_directory, abort, jsonify
import yt_dlp

app = Flask(__name__)

DOWNLOAD_FOLDER = os.path.join(os.getcwd(), 'downloads')
Path(DOWNLOAD_FOLDER).mkdir(parents=True, exist_ok=True)


def sanitize_filename(filename):
    """清理檔名，移除或替換掉URL和檔案系統中的不安全字元。"""
    # 移除URL錨點和查詢參數
    filename = re.split(r'[#?]', filename)[0]
    # 移除Windows/Linux/macOS中的非法檔名字符
    # 允許中文、日文、韓文字元，以及常見符號
    sanitized = re.sub(r'[\\/:*?"<>|]', '_', filename)
    return sanitized.strip()


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process_request():
    data = request.get_json()
    mode = data.get('mode')
    urls_text = data.get('urls')
    urls = [url.strip() for url in urls_text.strip().splitlines() if url.strip()]

    if not urls:
        return jsonify({'status': 'error', 'message': '請至少輸入一個網址'}), 400

    try:
        # --- 處理下載任務 ---
        # 注意：Web 下載不適合長時間、大量的批次任務。這裡我們仍以下載第一個有效的 URL 為例，
        # 但提供了一個更穩健的框架。批次下載更適合打包成 zip。
        if mode in ["multiple_videos", "playlist_videos"]:
            target_url = urls[0]

            ydl_opts = {
                'format': 'bestvideo[ext=mp4]+bestaudio[ext=m4a]/best[ext=mp4]/best',
                # 使用 yt-dlp 的 sanitize_filename 功能，並指定輸出範本
                'outtmpl': os.path.join(DOWNLOAD_FOLDER, '%(title)s - %(id)s.%(ext)s'),
                # 這個選項會讓 yt-dlp 盡可能清理檔名
                'restrictfilenames': True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(target_url, download=True)
                # prepare_filename 在 restrictfilenames=True 時會返回清理過的檔名
                filename = ydl.prepare_filename(info)
                result_filename = os.path.basename(filename)

            download_url = f"/files/{result_filename}"
            return jsonify({'status': 'success', 'download_url': download_url, 'filename': result_filename})

        # --- 【重大修正】處理資訊獲取任務，使其能處理多個 URL ---
        elif mode.startswith("info_"):
            all_data = []

            ydl_opts = {
                'quiet': True,
                'extract_flat': (mode == "info_playlist_fast"),
                'dump_single_json': True,
                'ignoreerrors': True,
            }

            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                # 循環處理所有傳入的 URL
                for url in urls:
                    try:
                        result = ydl.extract_info(url, download=False)
                        if not result: continue

                        # 處理播放清單和單一影片
                        if '_type' in result and result['_type'] == 'playlist':
                            entries = result.get('entries', [])
                            if entries: all_data.extend(filter(None, entries))
                        else:
                            all_data.append(result)
                    except Exception:
                        continue  # 如果單一 URL 失敗，就跳過繼續處理下一個

            if not all_data:
                return jsonify({'status': 'error', 'message': '無法獲取任何影片資訊。'}), 500

            # 後續的 Excel 處理邏輯保持不變
            processed_data = []
            for item in all_data:
                processed_data.append({
                    '標題': item.get('title', 'N/A'),
                    '觀看次數': item.get('view_count'),
                    '上傳時間': pd.to_datetime(item.get('upload_date', None), format='%Y%m%d',
                                               errors='coerce').strftime('%Y-%m-%d') if item.get(
                        'upload_date') else 'N/A',
                    '影片網址': item.get('webpage_url') or item.get('url', 'N/A')
                })

            df = pd.DataFrame(processed_data)
            result_filename = "youtube_info_export.xlsx"
            excel_filepath = os.path.join(DOWNLOAD_FOLDER, result_filename)
            df.to_excel(excel_filepath, index=False)

            download_url = f"/files/{result_filename}"
            return jsonify({'status': 'success', 'download_url': download_url, 'filename': result_filename,
                            'count': len(processed_data)})

    except Exception as e:
        return jsonify({'status': 'error', 'message': f"處理時發生嚴重錯誤: {e}"}), 500


# /files/<path:filename> 路由保持不變
@app.route('/files/<path:filename>')
def download_file(filename):
    try:
        return send_from_directory(DOWNLOAD_FOLDER, filename, as_attachment=True)
    except FileNotFoundError:
        abort(404)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)