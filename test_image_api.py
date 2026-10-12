import urllib.request
import urllib.parse
import json

headers = {'User-Agent': 'AcademicCurriculumAgent/1.0 (https://vfstr.ac.in; dean_academic@vfstr.ac.in)'}

topics = [
    ('Vibe coding', ['artificial intelligence code', 'machine learning programming', 'software programming']),
    ('Computer Architecture', ['microprocessor chip', 'integrated circuit', 'computer hardware']),
    ('Data Structures', ['binary tree algorithm', 'computer programming code', 'data structure']),
    ('Operating Systems', ['linux kernel operating system', 'computer terminal', 'operating system']),
    ('Database Management', ['database server', 'relational database', 'cloud database']),
    ('Computer Networks', ['computer network topology', 'network router server', 'telecommunication']),
    ('Cyber Security', ['cryptography computer security', 'network security', 'cyber security']),
    ('Web Development', ['html web design', 'web development code', 'internet programming']),
]

for topic, queries in topics:
    found = False
    for q in queries:
        q_enc = urllib.parse.quote(q)
        api_url = f"https://commons.wikimedia.org/w/api.php?action=query&generator=search&gsrnamespace=6&gsrsearch={q_enc}&gsrlimit=10&prop=imageinfo&iiprop=url|mime|size&format=json"
        req = urllib.request.Request(api_url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode())
                pages = data.get('query', {}).get('pages', {})
                for p in pages.values():
                    info = p.get('imageinfo', [{}])[0]
                    mime = info.get('mime', '')
                    width = info.get('width', 0)
                    if mime in ['image/jpeg', 'image/png'] and width >= 300:
                        img_url = info.get('url')
                        img_req = urllib.request.Request(img_url, headers=headers)
                        with urllib.request.urlopen(img_req, timeout=8) as r_img:
                            b = r_img.read()
                            print(f"[{topic}] -> SUCCESS! Found for query '{q}': {len(b)} bytes | {img_url.split('?')[0]}")
                        found = True
                        break
        except Exception as e:
            pass
        if found:
            break
    if not found:
        print(f"[{topic}] -> FAILED to find image")
