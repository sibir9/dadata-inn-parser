#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import requests
import pandas as pd
import time
import json
import re

# Ваш API-ключ Dadata
API_KEY = "4e539944b8c5498f253f1664df0ee01e7e0411ad"

def clean_inn(inn_str):
    """Очищает ИНН от мусора и приводит к строковому виду"""
    inn_str = str(inn_str).strip()
    # Убираем все нецифровые символы
    inn_clean = re.sub(r'\D', '', inn_str)
    return inn_clean

def get_company_data(inn):
    """Получение данных компании по ИНН через Dadata (расширенный эндпоинт)"""
    
    url = 'https://suggestions.dadata.ru/suggestions/api/4_1/rs/findById/party'
    
    headers = {
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Authorization': f'Token {API_KEY}'
    }
    
    data = {
        'query': inn,
        'branch_type': 'MAIN'
    }
    
    try:
        response = requests.post(url, json=data, headers=headers, timeout=10)
        
        if response.status_code == 200:
            result = response.json()
            suggestions = result.get('suggestions', [])
            
            if suggestions:
                org = suggestions[0].get('data', {})
                address = org.get('address', {}).get('value', '')
                
                # Получаем основной ОКВЭД и все дополнительные
                okved_main = org.get('okved', '')
                okveds_all = org.get('okveds', [])
                okveds_str = ', '.join(okveds_all) if okveds_all else ''
                
                # Пытаемся получить описание ОКВЭД (через отдельный запрос или оставляем пустым)
                # Dadata не возвращает описание, только код. Можно потом обогатить отдельно.
                
                return {
                    'ИНН': inn,
                    'Название_компании': org.get('name', {}).get('short_with_opf', ''),
                    'Название_полное': org.get('name', {}).get('full_with_opf', ''),
                    'ОКВЭД_основной': okved_main,
                    'ОКВЭД_все': okveds_str,
                    'Адрес': address,
                    'Статус': org.get('state', {}).get('status', ''),
                    'Тип': org.get('type', ''),  # LEGAL или INDIVIDUAL
                    'ОГРН': org.get('ogrn', ''),
                    'КПП': org.get('kpp', ''),
                    'Руководитель': org.get('management', {}).get('name', ''),
                    'Должность': org.get('management', {}).get('post', '')
                }
            else:
                return {
                    'ИНН': inn,
                    'Название_компании': 'НЕ НАЙДЕНО',
                    'ОКВЭД_основной': 'НЕ НАЙДЕН',
                    'Статус': 'НЕ НАЙДЕН'
                }
    except Exception as e:
        print(f'Ошибка для ИНН {inn}: {e}')
        return {
            'ИНН': inn,
            'Название_компании': f'ОШИБКА: {e}',
            'ОКВЭД_основной': 'ОШИБКА'
        }

def load_inns_from_file(file_path):
    """Загружает ИНН из текстового файла (по одному на строку)"""
    inns = []
    with open(file_path, 'r', encoding='utf-8') as f:
        for line in f:
            inn_clean = clean_inn(line)
            if inn_clean and len(inn_clean) >= 10:  # ИНН должен быть не короче 10 цифр
                inns.append(inn_clean)
    return inns

# ============ ОСНОВНОЙ СКРИПТ ============

print("=" * 60)
print("Парсер данных компаний по ИНН (Dadata API)")
print("=" * 60)

# Загружаем ИНН из файла
INPUT_FILE = 'inn.txt'
print(f"\n📂 Загрузка ИНН из файла: {INPUT_FILE}")

try:
    inns = load_inns_from_file(INPUT_FILE)
    print(f"✅ Найдено {len(inns)} уникальных ИНН")
except FileNotFoundError:
    print(f"❌ Файл {INPUT_FILE} не найден!")
    exit(1)

if not inns:
    print("❌ Не удалось загрузить ИНН. Проверьте файл.")
    exit(1)

# Обрабатываем каждый ИНН
results = []
total = len(inns)

print(f"\n🚀 Начинаем обработку {total} ИНН...")
print("⏳ Это может занять несколько минут...\n")

for i, inn in enumerate(inns, 1):
    print(f"  {i}/{total}: Обработка ИНН {inn}...", end=' ')
    
    data = get_company_data(inn)
    if data:
        results.append(data)
        print("✓")
    else:
        print("✗")
    
    # Пауза, чтобы не превысить лимит (10 000 запросов в день)
    time.sleep(0.2)

# Сохраняем результат в Excel
output_file = 'companies.xlsx'
df = pd.DataFrame(results)

# Переставляем колонки в нужном порядке
column_order = [
    'Название_компании',
    'ИНН',
    'ОКВЭД_основной',
    'ОКВЭД_все',
    'Адрес',
    'Статус',
    'Тип',
    'ОГРН',
    'КПП',
    'Руководитель',
    'Должность',
    'Название_полное'
]

# Оставляем только те колонки, которые есть в DataFrame
existing_columns = [col for col in column_order if col in df.columns]
df = df[existing_columns]

df.to_excel(output_file, index=False)

print("\n" + "=" * 60)
print(f"✅ Готово! Обработано {len(results)} из {total} ИНН")
print(f"📁 Результат сохранён в файл: {output_file}")
print("=" * 60)

# Выводим статистику по ОКВЭД
print("\n📊 Статистика по основным ОКВЭД (топ-20):")
okved_stats = df['ОКВЭД_основной'].value_counts().head(20)
for okved, count in okved_stats.items():
    print(f"  {okved}: {count}")

# Статистика по статусам
print("\n📊 Статусы компаний:")
status_stats = df['Статус'].value_counts()
for status, count in status_stats.items():
    print(f"  {status}: {count}")

# Статистика по типам
print("\n📊 Типы компаний:")
type_stats = df['Тип'].value_counts()
for t, count in type_stats.items():
    print(f"  {t}: {count}")
