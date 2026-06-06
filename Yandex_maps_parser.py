# Установка свежего Google Chrome (не старый Chromium из apt)
!pip install -q selenium pandas bs4 requests openpyxl gspread oauth2client
!wget -q https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
!apt-get install -q -y ./google-chrome-stable_current_amd64.deb

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def create_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    prefs = {'profile.managed_default_content_settings.images': 2}
    options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome(options=options)
    return driver

driver = create_driver()
print('Браузер успешно запущен!')
driver.quit()
print('Тест пройден — можно запускать скрапер!')

# Запуск в google.colab
from google.colab import auth
auth.authenticate_user()

# Если нужна выгрузка в гугл драйв
import gspread
from oauth2client.client import GoogleCredentials
from google.auth import default

creds, _ = default()
gc = gspread.authorize(creds)

sheet = gc.create("Tyumen Dentists")

worksheet = sheet.sheet1

worksheet.append_row(["Name", "Address", "Phone", "Website", "Email"])

# Сам скрабер

import re
import requests
import pickle
import os
from time import sleep, time
from datetime import datetime
from IPython.display import clear_output

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import pandas as pd

# ===================== НАСТРОЙКИ =====================
location = 'Тюмень'
title = 'стоматология'
count_of_units = 1000
CHECKPOINT_INTERVAL = 20

# === НОВЫЕ НАСТРОЙКИ ===
COLLECT_SOCIALS = True          # Собирать соцсети/мессенджеры
DEEP_SEARCH_IF_NO_SITE = True   # Искать сайт через Яндекс
USE_SMART_OCR = True            # Умный OCR (только при триггерах)
# =====================================================

# --- Установка Tesseract для OCR ---
if USE_SMART_OCR:
    print('📦 Установка Tesseract OCR...')
    !apt-get install -y tesseract-ocr tesseract-ocr-rus -qq
    !pip install -q pytesseract pillow
    import pytesseract
    from PIL import Image
    from io import BytesIO
    print('✅ Tesseract установлен')

# --- Прогресс-бар ---
def print_progress(current, total, start_time, stage='сбор данных', current_name=''):
    clear_output(wait=True)
    pct = current / total if total > 0 else 0
    bar_len = 40
    filled = int(bar_len * pct)
    bar = '█' * filled + '░' * (bar_len - filled)
    elapsed = time() - start_time
    if current > 0:
        eta = (elapsed / current) * (total - current)
        eta_str = f'{int(eta//60)}м {int(eta%60)}с'
    else:
        eta_str = '...'
    print('=' * 55)
    print(f'  🦷 {title.capitalize()} — {location}  |  {stage}')
    print('=' * 55)
    print(f'  [{bar}] {int(pct*100)}%')
    print(f'  📍 Обработано:  {current} / {total}')
    print(f'  ⏱  Прошло:      {int(elapsed//60)}м {int(elapsed%60)}с')
    print(f'  ⏳ Осталось:    {eta_str}')
    if current_name:
        print(f'  🔍 Сейчас:      {current_name[:45]}')
    print('=' * 55)

# --- Драйвер ---
def create_driver():
    options = Options()
    options.add_argument('--headless')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--disable-gpu')
    options.add_argument('--window-size=1920,1080')
    options.add_argument('--lang=ru-RU')
    options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    prefs = {
        'profile.managed_default_content_settings.images': 2,
        'intl.accept_languages': 'ru-RU,ru'
    }
    options.add_experimental_option('prefs', prefs)
    return webdriver.Chrome(options=options)

# --- Функция поиска страницы контактов ---
def find_contacts_page(soup, base_url):
    '''Ищет ссылку на страницу контактов'''
    keywords = ['контакт', 'contact', 'связь', 'связаться']
    
    for a in soup.find_all('a', href=True):
        href = a['href'].lower()
        text = a.get_text().lower()
        
        # Проверяем URL и текст ссылки
        if any(kw in href or kw in text for kw in keywords):
            # Формируем полный URL
            if href.startswith('http'):
                return href
            elif href.startswith('/'):
                return base_url.rstrip('/') + href
            else:
                return base_url.rstrip('/') + '/' + href
    return None

# --- Функция парсинга mailto ссылок ---
def parse_mailto_links(soup):
    '''Ищет mailto: ссылки - самый надёжный способ'''
    for a in soup.find_all('a', href=True):
        if a['href'].startswith('mailto:'):
            email = a['href'].replace('mailto:', '').split('?')[0]  # убираем параметры
            # Валидация
            if re.match(r'^[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}$', email):
                return email
    return None

# --- Улучшенный парсинг email из текста ---
def parse_text_for_email(text, exclude_images=True):
    '''Парсит email из текста с фильтрацией'''
    found = re.findall(r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}', text)
    
    if exclude_images:
        # Фильтруем файлы изображений и другие ложные срабатывания
        blacklist_extensions = ['.png', '.jpg', '.jpeg', '.gif', '.svg', '.webp', '.ico', '.css', '.js']
        valid_emails = [e for e in found if not any(e.lower().endswith(ext) for ext in blacklist_extensions)]
    else:
        valid_emails = found
    
    # Исключаем популярные ложные срабатывания
    spam_patterns = ['rating@mail.ru', 'example@', 'test@', 'noreply@', 'no-reply@']
    valid_emails = [e for e in valid_emails if not any(sp in e.lower() for sp in spam_patterns)]
    
    return valid_emails[0] if valid_emails else None

# --- Функция извлечения соцсетей ---
def extract_socials(soup, url_text=''):
    socials = {
        'instagram': 'null',
        'vk': 'null',
        'facebook': 'null',
        'telegram': 'null',
        'whatsapp': 'null'
    }
    
    # Паттерны социальных сетей
    patterns = {
        'instagram': r'instagram\.com/([a-zA-Z0-9_.]+)',
        'vk': r'vk\.com/([a-zA-Z0-9_]+)',
        'facebook': r'facebook\.com/([a-zA-Z0-9.]+)',
        'telegram': r't\.me/([a-zA-Z0-9_]+)',
        'whatsapp': r'(wa\.me/[0-9]+|api\.whatsapp\.com/send\?phone=[0-9]+)'
    }
    
    # Ищем в ссылках
    for a in soup.find_all('a', href=True):
        href = a['href'].lower()
        for social, pattern in patterns.items():
            if socials[social] == 'null':
                match = re.search(pattern, href)
                if match:
                    socials[social] = match.group(0)
    
    # Дополнительно ищем в тексте страницы
    if url_text:
        for social, pattern in patterns.items():
            if socials[social] == 'null':
                match = re.search(pattern, url_text.lower())
                if match:
                    socials[social] = match.group(0)
    
    return socials

# --- Функция поиска сайта через Яндекс ---
def find_site_via_yandex(name, address, driver):
    try:
        query = f'{name} {address} сайт'
        search_url = f'https://yandex.ru/search/?text={query.replace(" ", "+")}'
        driver.get(search_url)
        sleep(2)
        
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        
        # Ищем первую органическую ссылку
        for link in soup.find_all('a', href=True):
            href = link['href']
            # Пропускаем служебные ссылки Яндекса
            if any(skip in href for skip in ['yandex.', 'ya.ru', 'wikipedia', 'google', 'vk.com']):
                continue
            # Берём первую подходящую
            if href.startswith('http'):
                return href
        return None
    except:
        return None

# --- Умный OCR (только при триггерах) ---
def smart_ocr_email(soup, site_url, page_text=''):
    if not USE_SMART_OCR:
        return None, None
    
    triggers = []
    img_urls = []
    
    # Триггер 1: JavaScript с @
    scripts = soup.find_all('script')
    for script in scripts:
        if script.string and '@' in script.string:
            triggers.append('js-protection')
            break
    
    # Триггер 2: Текст "email:" или "почта:" рядом с картинкой
    if any(kw in page_text.lower() for kw in ['электронная почта:', 'email:', 'e-mail:', 'почта:']):
        triggers.append('email-label-found')
    
    # Триггер 3: Картинки с ключевыми словами
    for img in soup.find_all('img'):
        src = img.get('src', '').lower()
        alt = img.get('alt', '').lower()
        title = img.get('title', '').lower()
        
        if any(kw in src+alt+title for kw in ['email', 'mail', 'contact', 'контакт', '@', 'pochta']):
            triggers.append('image-contact')
            # Формируем полный URL картинки
            full_src = img.get('src', '')
            if full_src:
                if full_src.startswith('http'):
                    img_urls.append(full_src)
                elif full_src.startswith('/'):
                    base_url = '/'.join(site_url.split('/')[:3])
                    img_urls.append(base_url + full_src)
                elif not full_src.startswith('#'):
                    base_url = '/'.join(site_url.split('/')[:-1])
                    img_urls.append(base_url + '/' + full_src)
    
    if not triggers:
        return None, None
    
    # Запускаем OCR
    try:
        for img_url in img_urls[:3]:  # максимум 3 картинки
            try:
                response = requests.get(img_url, timeout=5, headers={'User-Agent': 'Mozilla/5.0'})
                img = Image.open(BytesIO(response.content))
                text = pytesseract.image_to_string(img, lang='rus+eng')
                
                # Ищем email в распознанном тексте
                email = parse_text_for_email(text, exclude_images=True)
                if email:
                    return email, 'ocr-image'
            except:
                continue
        
        return None, None
    except:
        return None, None

# --- Проверка кэша ссылок ---
CACHE_FILE = 'links_cache.pkl'
href_list = None

if os.path.exists(CACHE_FILE):
    print('🔄 Найден кэш ссылок!')
    with open(CACHE_FILE, 'rb') as f:
        href_list = pickle.load(f)
    print(f'✅ Загружено {len(href_list)} ссылок из кэша')

# --- Скроллинг (если кэша нет) ---
if href_list is None:
    driver = create_driver()
    query = f"{title} {location}"
    url = f"https://yandex.ru/maps/?text={query.replace(' ', '%20')}&lang=ru_RU"
    driver.get(url)
    sleep(5)

    print('⏳ Начинаем скроллинг Яндекс Карт...')
    n = 0
    scroll_start = time()
    while True:
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        org_links = set()
        for a in soup.find_all('a', href=True):
            href = a['href']
            if '/org/' in href or '/maps/org' in href:
                if href.startswith('/'):
                    href = 'https://yandex.ru' + href
                base = re.match(r'(https://yandex\.ru/maps/org/[^/]+/\d+)', href)
                if base:
                    org_links.add(base.group(1))

        elapsed = int(time() - scroll_start)
        try:
            cards = driver.find_elements(By.CSS_SELECTOR, '[class*="search-business-snippet-view"]')
            cards_count = len(cards)
        except:
            cards_count = 0
        
        clear_output(wait=True)
        print('=' * 55)
        print(f'  🗺️  Скроллинг Яндекс Карт')
        print('=' * 55)
        print(f'  🔍 Найдено организаций:  {len(org_links)}')
        print(f'  📄 Карточек на странице:  {cards_count}')
        print(f'  ⏱  Прошло времени:        {elapsed}с')
        print('=' * 55)

        if len(org_links) >= count_of_units:
            break

        try:
            if cards_count > 0:
                driver.execute_script('arguments[0].scrollIntoView(true);', cards[-1])
                sleep(1.5)
                n_new = cards_count
                if n_new == n:
                    n += 1
                    if n >= 10:
                        break
                else:
                    n = 0
                n = n_new
        except:
            break

    href_list = list(org_links)
    driver.quit()

    with open(CACHE_FILE, 'wb') as f:
        pickle.dump(href_list, f)
    print(f'\n✅ Собрано {len(href_list)} ссылок | Кэш сохранён')

# --- Проверка чекпоинтов ---
checkpoint_files = sorted([f for f in os.listdir('/content') if f.startswith('checkpoint_') and f.endswith('.pkl')])
start_idx = 1
keys = {
    'href': [], 'name': [], 'adress': [], 'phone': [],
    'rate': [], 'rate_count': [], 'site': [], 'site_source': [], 
    'email': [], 'email_source': [],
    'instagram': [], 'vk': [], 'facebook': [], 'telegram': [], 'whatsapp': [],
    'average_bill': []
}

if checkpoint_files:
    latest_checkpoint = checkpoint_files[-1]
    print(f'🔄 Найден чекпоинт: {latest_checkpoint}')
    with open(f'/content/{latest_checkpoint}', 'rb') as f:
        keys = pickle.load(f)
    start_idx = len(keys['href']) + 1
    print(f'✅ Продолжаем с записи #{start_idx}')

# --- Сбор данных ---
driver = create_driver()
start_time = time()

for idx in range(start_idx, len(href_list) + 1):
    i = href_list[idx - 1]
    name_hint = i.split('/')[-2] if '/' in i else i
    print_progress(idx, len(href_list), start_time, 'сбор данных', name_hint)

    driver.get(i)
    sleep(2)
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    keys['href'].append(i)

    # Название
    try:
        name = soup.find('h1', class_='orgpage-header-view__header').text
        keys['name'].append(name)
    except:
        name = 'null'
        keys['name'].append('null')

    # Адрес
    try:
        address = soup.find('a', class_='orgpage-header-view__address').text
        keys['adress'].append(address)
    except:
        address = 'null'
        keys['adress'].append('null')

    # Телефон
    try:
        keys['phone'].append(soup.find('div', class_='orgpage-phones-view__phone-number').text)
    except:
        keys['phone'].append('null')

    # Рейтинг
    try:
        keys['rate'].append(soup.find('span', class_='business-rating-badge-view__rating-text').text)
    except:
        keys['rate'].append('null')

    # Кол-во отзывов
    try:
        keys['rate_count'].append(soup.find('div', class_='business-header-rating-view__text _clickable').text)
    except:
        keys['rate_count'].append('null')

    # === СБОР СОЦСЕТЕЙ С ЯНДЕКС КАРТ ===
    if COLLECT_SOCIALS:
        socials_yandex = extract_socials(soup, driver.page_source)
    else:
        socials_yandex = {'instagram': 'null', 'vk': 'null', 'facebook': 'null', 'telegram': 'null', 'whatsapp': 'null'}

    # Сайт
    site_source = 'null'
    try:
        site_text = soup.find('span', class_='business-urls-view__text').text
        keys['site'].append(site_text)
        site_source = 'yandex_maps'
    except:
        site_text = 'null'
        keys['site'].append('null')
        
        # === ГЛУБОКИЙ ПОИСК САЙТА ===
        if DEEP_SEARCH_IF_NO_SITE and name != 'null' and address != 'null':
            found_site = find_site_via_yandex(name, address, driver)
            if found_site:
                site_text = found_site
                keys['site'][-1] = found_site
                site_source = 'yandex_search'
    
    keys['site_source'].append(site_source)

    # === ПАРСИНГ САЙТА С ПРИОРИТЕТНЫМ ПОИСКОМ EMAIL ===
    email = 'null'
    email_source = 'null'
    socials_site = {'instagram': 'null', 'vk': 'null', 'facebook': 'null', 'telegram': 'null', 'whatsapp': 'null'}
    
    if site_text != 'null':
        try:
            site_url = site_text if site_text.startswith('http') else 'http://' + site_text
            response = requests.get(
                site_url,
                timeout=(3, 5),
                headers={'User-Agent': 'Mozilla/5.0'},
                allow_redirects=True
            )
            site_soup = BeautifulSoup(response.text, 'html.parser')
            base_url = '/'.join(site_url.split('/')[:3])
            
            # === ПРИОРИТЕТ 1: mailto: ссылки (самое надёжное) ===
            email = parse_mailto_links(site_soup)
            if email:
                email_source = 'mailto'
            
            # === ПРИОРИТЕТ 2: Страница "Контакты" ===
            if not email:
                contacts_url = find_contacts_page(site_soup, base_url)
                if contacts_url:
                    try:
                        contacts_resp = requests.get(contacts_url, timeout=(3, 5), headers={'User-Agent': 'Mozilla/5.0'})
                        contacts_soup = BeautifulSoup(contacts_resp.text, 'html.parser')
                        
                        # Сначала ищем mailto на странице контактов
                        email = parse_mailto_links(contacts_soup)
                        if email:
                            email_source = 'contacts_mailto'
                        else:
                            # Потом парсим текст
                            email = parse_text_for_email(contacts_resp.text)
                            if email:
                                email_source = 'contacts_page'
                    except:
                        pass
            
            # === ПРИОРИТЕТ 3: Парсинг главной страницы ===
            if not email:
                email = parse_text_for_email(response.text)
                if email:
                    email_source = 'website'
            
            # === ПРИОРИТЕТ 4: Умный OCR (если триггеры) ===
            if not email:
                ocr_email, ocr_source = smart_ocr_email(site_soup, site_url, response.text)
                if ocr_email:
                    email = ocr_email
                    email_source = ocr_source
            
            # Собираем соцсети с сайта
            if COLLECT_SOCIALS:
                socials_site = extract_socials(site_soup, response.text)
        except:
            pass
    
    keys['email'].append(email)
    keys['email_source'].append(email_source)
    
    # Объединяем соцсети (приоритет: Яндекс Карты, потом сайт)
    for social in ['instagram', 'vk', 'facebook', 'telegram', 'whatsapp']:
        if socials_yandex[social] != 'null':
            keys[social].append(socials_yandex[social])
        elif socials_site[social] != 'null':
            keys[social].append(socials_site[social])
        else:
            keys[social].append('null')

    # Средний чек
    try:
        keys['average_bill'].append(soup.find('span', class_='business-features-view__valued-value').text)
    except:
        keys['average_bill'].append('null')

    # Чекпоинт
    if idx % CHECKPOINT_INTERVAL == 0:
        with open(f'/content/checkpoint_{idx}.pkl', 'wb') as f:
            pickle.dump(keys, f)

driver.quit()

# Удаляем чекпоинты
for f in checkpoint_files:
    os.remove(f'/content/{f}')
if os.path.exists(CACHE_FILE):
    os.remove(CACHE_FILE)

# Создаём DataFrame
now = datetime.now()
df = pd.DataFrame(keys)

# Сохраняем Excel
filename = f'{location}-{title}-{now.strftime("%Y-%m-%d_%H-%M-%S")}.xlsx'
df.to_excel(filename, index=False)

# Статистика
clear_output(wait=True)
print('=' * 60)
print(f'  ✅ ГОТОВО!')
print('=' * 60)
print(f'  📊 Собрано записей: {len(df)}')
print(f'  📁 Excel файл:      {filename}')
print(f'  📞 С телефоном:     {(df["phone"] != "null").sum()}')
print(f'  🌐 С сайтом:        {(df["site"] != "null").sum()}')
print(f'  📧 С email:         {(df["email"] != "null").sum()}')
if USE_SMART_OCR:
    print(f'     └─ OCR:          {(df["email_source"].str.contains("ocr")).sum()}')
print(f'\n📍 Источники сайтов:')
print(f'   Яндекс Карты:     {(df["site_source"] == "yandex_maps").sum()}')
print(f'   Яндекс Поиск:     {(df["site_source"] == "yandex_search").sum()}')
if COLLECT_SOCIALS:
    print(f'  📱 Instagram:       {(df["instagram"] != "null").sum()}')
    print(f'  📱 VK:              {(df["vk"] != "null").sum()}')
    print(f'  📱 Telegram:        {(df["telegram"] != "null").sum()}')
    print(f'  📱 WhatsApp:        {(df["whatsapp"] != "null").sum()}')
print('=' * 60)

# Скачиваем
from google.colab import files
files.download(filename)

print('\n📥 Файл загружается в браузер...')
print('\nПервые 10 записей:')
print(df.head(10))
