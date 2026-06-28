import urllib.request, re, os

url = 'https://www.aigei.com/item/nai_long_dong_14.html'
req = urllib.request.Request(url, headers={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
})
resp = urllib.request.urlopen(req, timeout=10)
html = resp.read().decode('utf-8', errors='replace')

image_urls = re.findall(r'https?://[^"\']+\.(?:jpg|jpeg|png|gif|webp)(?:\?[^"\'\\s]*)?', html)
for img in image_urls:
    if 'logo' not in img.lower() and 'icon' not in img.lower():
        print(f"Found: {img}")
