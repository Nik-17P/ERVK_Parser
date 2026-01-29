<img width="3025" height="201" alt="image" src="https://github.com/user-attachments/assets/c84373d9-e07d-4484-b124-fc422df2e241" />

Пример выгрузки в эксель

Запускаете скрипт - у вас автоматически открывается Google Chrome браузер, сайт
Открылся? Если нет, то ставим пакет

pip install playwright pandas openpyxl
python -m playwright install

И не забываем убедиться, что другие библиотеки установлены. Всего понадобится:
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import pandas as pd
import time
import datetime
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
import json
from webdriver_manager.chrome import ChromeDriverManager
import glob

Если все открылось, то следуем инструкциям в терминале - выставляем фильтры
Фильтры поставили - в терминале кликаем Enter
И начнется сбор информации по всем страницам

Да, это долго
