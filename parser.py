import os
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import re
from datetime import datetime


class OzonSellerParser:
    def __init__(self, seller_urls, output_folder='prices_with_co-investment'):
        """
        Инициализация парсера
        
        Args:
            seller_urls: Список URL продавцов на Ozon или путь к файлу с URL
            output_folder: Папка для сохранения Excel файла
        """
        self.seller_urls = self._parse_input_urls(seller_urls)
        self.output_folder = output_folder
        self.driver = None
        self.products_data = []
        self.visited_urls = set()
        
    def _parse_input_urls(self, input_data):
        """
        Парсинг входных данных для получения списка URL
        """
        urls = []
        
        if isinstance(input_data, str):
            if input_data.endswith('.txt') and os.path.exists(input_data):
                print(f"Чтение URL из файла: {input_data}")
                with open(input_data, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#'):
                            urls.append(line)
                print(f"Найдено URL в файле: {len(urls)}")
            else:
                urls.append(input_data)
        
        elif isinstance(input_data, list):
            urls = input_data
        
        valid_urls = []
        for url in urls:
            if url.startswith('http') and 'ozon.ru/seller/' in url:
                valid_urls.append(url)
            else:
                print(f"⚠ Пропущен некорректный URL: {url}")
        
        if not valid_urls:
            raise ValueError("Не найдено корректных URL продавцов Ozon")
        
        return valid_urls
    
    def setup_driver(self):
        """Настройка драйвера Selenium"""
        options = webdriver.ChromeOptions()
        
        # Основные настройки
        options.add_argument('--start-maximized')
        options.add_argument('--disable-blink-features=AutomationControlled')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-gpu')
        
        # Реалистичный User-Agent
        options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # Отключаем автоматизацию
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=options)
        
        # Выполняем скрипты для сокрытия автоматизации
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    def check_and_solve_captcha(self, require_solution=True):
        """
        Проверка наличия капчи и запрос на ручное решение
        Возвращает True если капча решена или не найдена, False если ошибка
        
        Args:
            require_solution: Если True - требует решения капчи, если False - только проверяет
        """
        print("Проверка на наличие капчи...")
        
        # 1. Проверяем по заголовку страницы - самый надежный признак
        try:
            title = self.driver.title.lower()
            if 'antibot captcha' in title or 'капча' in title:
                print("⚠ Обнаружена капча по заголовку страницы")
                return self._handle_captcha_page(require_solution)
        except:
            pass
        
        # 2. Проверяем структуру страницы капчи Ozon
        captcha_indicators = [
            # Точные ID элементов капчи Ozon
            ("//div[@id='captcha-container']", "контейнер капчи Ozon"),
            ("//div[@id='captcha']", "виджет капчи Ozon"),
            ("//div[@id='slider-background']", "слайдер капчи Ozon"),
            
            # Точные классы капчи Ozon
            ("//div[contains(@class, 'captcha-container') and @id='captcha-container']", "контейнер капчи"),
            
            # Точные тексты на странице капчи (только полные фразы)
            ("//*[text()='Подтвердите, что вы не бот']", "текст подтверждения капчи"),
            ("//*[contains(text(), 'Передвиньте ползунок, чтобы пазл попал в контур')]", "инструкция капчи"),
            ("//*[contains(text(), 'Подтвердите, что вы не робот') and contains(@class, 'title')]", "заголовок капчи"),
        ]
        
        captcha_found = False
        captcha_type = ""
        
        # Проверяем точные индикаторы
        for xpath, description in captcha_indicators:
            try:
                elements = self.driver.find_elements(By.XPATH, xpath)
                for element in elements:
                    try:
                        if element.is_displayed():
                            # Для текстовых элементов проверяем контекст
                            if 'text()' in xpath or 'contains(text()' in xpath:
                                # Проверяем, что элемент находится в контексте капчи
                                parent_html = element.get_attribute('outerHTML')
                                if 'captcha' not in parent_html.lower() and 'bot' not in parent_html.lower():
                                    # Пропускаем, если нет контекста капчи
                                    continue
                            
                            captcha_found = True
                            captcha_type = description
                            print(f"⚠ Обнаружена капча: {description}")
                            break
                    except:
                        continue
                if captcha_found:
                    break
            except:
                continue
        
        # 3. Проверяем iframe reCAPTCHA (Google)
        if not captcha_found:
            try:
                iframes = self.driver.find_elements(By.TAG_NAME, "iframe")
                for iframe in iframes:
                    try:
                        src = iframe.get_attribute('src')
                        if src and 'google.com/recaptcha' in src:
                            captcha_found = True
                            captcha_type = "Google reCAPTCHA iframe"
                            print(f"⚠ Обнаружена капча: {captcha_type}")
                            break
                    except:
                        continue
            except:
                pass
        
        # 4. Проверяем скрытые поля капчи Ozon
        if not captcha_found:
            try:
                hidden_inputs = ["captcha-input", "incident", "complaints-token", "captcha-ip", "captcha-date"]
                for input_id in hidden_inputs:
                    element = self.driver.find_element(By.ID, input_id)
                    if element:
                        captcha_found = True
                        captcha_type = f"скрытое поле капчи ({input_id})"
                        print(f"⚠ Обнаружена капча: {captcha_type}")
                        break
            except:
                pass
        
        # 5. Проверяем URL и структуру DOM как страницу капчи
        if not captcha_found:
            try:
                current_url = self.driver.current_url.lower()
                page_source = self.driver.page_source.lower()
                
                # Если на странице есть много элементов капчи
                if ('captcha' in page_source and 'slider' in page_source and 
                    'puzzle' in page_source and 'antibot' in page_source):
                    captcha_found = True
                    captcha_type = "страница капчи по структуре DOM"
                    print(f"⚠ Обнаружена капча: {captcha_type}")
            except:
                pass
        
        if captcha_found:
            print(f"\n{'='*60}")
            print("ОБНАРУЖЕНА КАПЧА!")
            print(f"Тип: {captcha_type}")
            print(f"{'='*60}")
            
            return self._handle_captcha_page(require_solution)
        
        print("✓ Капча не обнаружена")
        return True
    
    def _handle_captcha_page(self, require_solution):
        """Обработка страницы с капчей"""
        # Делаем скриншот капчи для наглядности
        try:
            self.create_output_folder()
            screenshot_path = os.path.join(self.output_folder, "captcha_screenshot.png")
            self.driver.save_screenshot(screenshot_path)
            print(f"Скриншот сохранен: {screenshot_path}")
        except:
            pass
        
        if require_solution:
            print("\nИнструкция по решению капчи:")
            print("1. Посмотрите на страницу браузера")
            print("2. Найдите капчу (обычно слайдер с пазлом)")
            print("3. Перетащите ползунок, чтобы собрать пазл")
            print("4. После решения нажмите Enter в этом окне")
            print("\n" + "="*60)
            
            # Ожидаем ручного решения
            input("Решите капчу в браузере и нажмите Enter для продолжения...")
            
            # Ждем немного после решения
            time.sleep(3)
            
            # Проверяем, осталась ли капча
            if self.check_and_solve_captcha(require_solution=False):
                print("✓ Капча решена успешно!")
                return True
            else:
                print("⚠ Капча все еще присутствует. Попробуйте еще раз.")
                return self._handle_captcha_page(require_solution)
        else:
            print("⚠ Капча обнаружена, но решение не требуется в данном контексте")
            return True
    
    def safe_get(self, url, max_retries=3):
        """
        Безопасный переход по URL с обработкой капчи
        """
        for attempt in range(max_retries):
            try:
                print(f"Переход по URL (попытка {attempt + 1}/{max_retries}): {url[:80]}...")
                self.driver.get(url)
                
                # Ждем загрузки страницы
                time.sleep(random.uniform(2, 4))
                
                # Проверяем капчу с требованием решения
                if not self.check_and_solve_captcha(require_solution=True):
                    print("⚠ Не удалось решить капчу, пробуем еще раз...")
                    continue
                
                # Проверяем, что страница загрузилась
                try:
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.TAG_NAME, "body"))
                    )
                    return True
                except:
                    print(f"Страница не загрузилась полностью, пробуем еще раз...")
                    continue
                    
            except Exception as e:
                print(f"Ошибка при загрузке страницы: {e}")
                time.sleep(random.uniform(3, 5))
        
        print(f"Не удалось загрузить страницу после {max_retries} попыток")
        return False
    
    def human_like_scroll(self, scroll_distance=None, scroll_duration=1.0):
        """Имитация человеческой прокрутки"""
        if scroll_distance is None:
            self.driver.execute_script("window.scrollTo({top: document.body.scrollHeight, behavior: 'smooth'});")
        else:
            self.driver.execute_script(f"""
                window.scrollBy({{
                    top: {scroll_distance},
                    behavior: 'smooth'
                }});
            """)
        
        time.sleep(scroll_duration + random.uniform(0.2, 0.5))
    
    def human_like_pause(self, min_time=0.5, max_time=2.0):
        """Случайная пауза как у человека"""
        pause_time = random.uniform(min_time, max_time)
        time.sleep(pause_time)
    
    def get_total_products_count(self):
        """Получение общего количества товаров продавца"""
        print("Определение общего количества товаров...")
        
        total_count = 0
        
        try:
            count_selectors = [
                "//*[contains(text(), 'товар') and contains(text(), 'найдено')]",
                "//*[contains(text(), 'товар') and contains(text(), 'всего')]",
                "//*[contains(@class, 'total') and contains(text(), 'товар')]",
            ]
            
            for selector in count_selectors:
                try:
                    elements = self.driver.find_elements(By.XPATH, selector)
                    for element in elements:
                        text = element.text.strip()
                        numbers = re.findall(r'\d+', text)
                        if numbers:
                            potential_count = max(map(int, numbers))
                            if potential_count > total_count and potential_count < 1000:
                                total_count = potential_count
                except:
                    continue
        except:
            pass
        
        if total_count == 0:
            try:
                product_elements = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/product/']")
                total_count = len(product_elements)
            except:
                pass
        
        return total_count
    
    def load_all_products_humanlike(self, seller_name):
        """Загрузка ВСЕХ товаров с имитацией человеческого поведения"""
        print(f"\n" + "="*60)
        print(f"ЗАГРУЗКА ТОВАРОВ ПРОДАВЦА: {seller_name}")
        print("="*60)
        
        # Проверяем капчу перед началом загрузки
        if not self.check_and_solve_captcha(require_solution=True):
            print("⚠ Не удалось решить капчу, пропускаем этого продавца")
            return []
        
        all_product_urls = set()
        last_url_count = 0
        same_count_iterations = 0
        max_same_count = 8
        iteration = 0
        max_iterations = 100
        
        total_expected = self.get_total_products_count()
        if total_expected > 0:
            print(f"Ожидаемое количество товаров: {total_expected}")
        
        print("\nНачинаем загрузку товаров...")
        
        while iteration < max_iterations and same_count_iterations < max_same_count:
            iteration += 1
            print(f"\nИтерация {iteration}...")
            
            # Плавная прокрутка вниз
            self.human_like_scroll()
            self.human_like_pause(1.0, 2.0)
            
            # Периодически проверяем капчу (реже, чтобы не мешать)
            if iteration % 15 == 0:
                # Проверяем только по заголовку и структуре, не по текстам
                try:
                    title = self.driver.title.lower()
                    if 'antibot captcha' in title or 'капча' in title:
                        print("⚠ Обнаружена капча во время загрузки!")
                        if not self.check_and_solve_captcha(require_solution=True):
                            print("⚠ Не удалось решить капчу, прекращаем загрузку")
                            break
                except:
                    pass
            
            # Случайная прокрутка вверх-вниз
            if random.random() < 0.3:
                self.human_like_scroll(-random.randint(300, 800))
                self.human_like_pause(0.5, 1.5)
                self.human_like_scroll(random.randint(400, 900))
                self.human_like_pause(0.5, 1.5)
            
            # Сбор товаров
            try:
                WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a[href*='/product/']"))
                )
                
                product_links = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/product/']")
                new_urls = set()
                
                for link in product_links:
                    try:
                        href = link.get_attribute('href')
                        if href and '/product/' in href:
                            clean_url = href.split('?')[0].split('#')[0]
                            if 'ozon.ru' in clean_url:
                                new_urls.add(clean_url)
                    except:
                        continue
                
                before_count = len(all_product_urls)
                all_product_urls.update(new_urls)
                after_count = len(all_product_urls)
                new_items = after_count - before_count
                
                print(f"  Найдено ссылок: {len(product_links)}")
                print(f"  Уникальных товаров: {after_count}")
                
                if new_items > 0:
                    print(f"  Новых товаров: +{new_items}")
                    same_count_iterations = 0
                    
                    if total_expected > 0 and after_count >= total_expected * 0.95:
                        print(f"  ✓ Загружено {after_count} из {total_expected} товаров (95%+)")
                        same_count_iterations = max_same_count - 1
                else:
                    same_count_iterations += 1
                    print(f"  Новых товаров нет ({same_count_iterations}/{max_same_count})")
                    
                    if same_count_iterations == 2:
                        self.try_click_show_more()
                    
                    if same_count_iterations == 4:
                        self.find_hidden_products()
                
            except Exception as e:
                print(f"  Ошибка при сборе товаров: {e}")
                same_count_iterations += 1
            
            if len(all_product_urls) >= 170:
                print(f"  ✓ Загружено {len(all_product_urls)} товаров (близко к цели)")
                if same_count_iterations >= 3:
                    break
            
            self.human_like_pause(1.5, 3.0)
        
        # Финальный сбор
        print("\nФинальный сбор всех товаров...")
        
        for i in range(3):
            self.human_like_scroll()
            self.human_like_pause(1.0, 2.0)
        
        final_urls = self.collect_all_product_urls()
        all_product_urls.update(final_urls)
        
        print(f"\n✓ Загрузка завершена")
        print(f"✓ Всего собрано уникальных товаров: {len(all_product_urls)}")
        
        if total_expected > 0:
            percentage = len(all_product_urls) / total_expected * 100
            print(f"✓ Прогресс: {percentage:.1f}% от ожидаемых {total_expected} товаров")
        
        return list(all_product_urls)
    
    def try_click_show_more(self):
        """Попытка кликнуть на кнопку 'Показать ещё'"""
        print("  Пробуем найти кнопку 'Показать ещё'...")
        
        show_more_selectors = [
            "//button[contains(., 'Показать ещё')]",
            "//button[contains(., 'Показать еще')]",
            "//div[contains(., 'Показать ещё') and @role='button']",
            "//button[contains(@class, 'show-more')]",
            "//button[@data-widget='showMore']",
        ]
        
        for selector in show_more_selectors:
            try:
                elements = self.driver.find_elements(By.XPATH, selector)
                for element in elements:
                    try:
                        if element.is_displayed():
                            self.driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                            self.human_like_pause(0.5, 1.0)
                            element.click()
                            print(f"    ✓ Нажата кнопка: {selector}")
                            self.human_like_pause(2.0, 3.0)
                            return True
                    except:
                        try:
                            self.driver.execute_script("arguments[0].click();", element)
                            print(f"    ✓ Нажата кнопка через JS: {selector}")
                            self.human_like_pause(2.0, 3.0)
                            return True
                        except:
                            continue
            except:
                continue
        
        print("    Кнопка 'Показать ещё' не найдена")
        return False
    
    def find_hidden_products(self):
        """Поиск скрытых товаров"""
        print("  Поиск скрытых товаров...")
        
        try:
            self.driver.execute_script("""
                document.querySelectorAll('img[loading="lazy"]').forEach(img => {
                    if (img.dataset.src) {
                        img.src = img.dataset.src;
                    }
                });
            """)
            print("    Загружены ленивые изображения")
            self.human_like_pause(1.0, 2.0)
        except:
            pass
    
    def collect_all_product_urls(self):
        """Сбор всех URL товаров"""
        all_urls = set()
        
        try:
            product_elements = self.driver.find_elements(By.CSS_SELECTOR, "a[href*='/product/']")
            for element in product_elements:
                try:
                    href = element.get_attribute('href')
                    if href and '/product/' in href:
                        clean_url = href.split('?')[0].split('#')[0]
                        if 'ozon.ru' in clean_url:
                            all_urls.add(clean_url)
                except:
                    continue
        except:
            pass
        
        additional_selectors = [
            "//a[contains(@href, '/product/')]",
            "//a[starts-with(@href, 'https://www.ozon.ru/product/')]",
        ]
        
        for selector in additional_selectors:
            try:
                elements = self.driver.find_elements(By.XPATH, selector)
                for element in elements:
                    try:
                        href = element.get_attribute('href')
                        if href and '/product/' in href:
                            clean_url = href.split('?')[0].split('#')[0]
                            if 'ozon.ru' in clean_url:
                                all_urls.add(clean_url)
                    except:
                        continue
            except:
                continue
        
        return all_urls
    
    def parse_product_page(self, product_url):
        """Парсинг данных с карточки товара"""
        if product_url in self.visited_urls:
            return
        
        print(f"Парсинг: {product_url[:70]}...")
        
        try:
            # Используем безопасный переход с обработкой капчи
            if not self.safe_get(product_url):
                print(f"  ⚠ Не удалось загрузить страницу товара")
                self.products_data.append({'sku': '', 'name': '', 'price': '', 'seller_url': self.current_seller_url})
                return
            
            self.visited_urls.add(product_url)
            
            # Проверяем капчу на странице товара с требованием решения
            if not self.check_and_solve_captcha(require_solution=True):
                print(f"  ⚠ Не удалось решить капчу на странице товара")
                self.products_data.append({'sku': '', 'name': '', 'price': '', 'seller_url': self.current_seller_url})
                return
            
            product_data = {'sku': '', 'name': '', 'price': '', 'seller_url': self.current_seller_url}
            
            # Артикул
            try:
                sku_elements = self.driver.find_elements(By.CSS_SELECTOR, "div.ga5_3_11-a2.tsBodyControl400Small")
                for element in sku_elements:
                    try:
                        text = element.text.strip()
                        if 'Артикул' in text:
                            numbers = re.findall(r'\d+', text)
                            if numbers:
                                product_data['sku'] = numbers[-1]
                                break
                    except:
                        continue
                
                if not product_data['sku']:
                    sku_elements = self.driver.find_elements(By.XPATH, "//*[contains(., 'Артикул')]")
                    for element in sku_elements:
                        try:
                            text = element.text.strip()
                            if 'Артикул' in text:
                                numbers = re.findall(r'\d+', text)
                                if numbers:
                                    product_data['sku'] = numbers[-1]
                                    break
                        except:
                            continue
            except:
                pass
            
            # Название
            try:
                name_elements = self.driver.find_elements(By.CSS_SELECTOR, "h1.pdp_gb9.tsHeadline550Medium")
                for element in name_elements:
                    try:
                        name = element.text.strip()
                        if name:
                            product_data['name'] = name
                            break
                    except:
                        continue
                
                if not product_data['name']:
                    name_elements = self.driver.find_elements(By.CSS_SELECTOR, "h1")
                    for element in name_elements:
                        try:
                            name = element.text.strip()
                            if name and len(name) > 3:
                                product_data['name'] = name
                                break
                        except:
                            continue
            except:
                pass
            
            # Цена
            try:
                price_elements = self.driver.find_elements(By.CSS_SELECTOR, "span.tsHeadline600Large")
                for element in price_elements:
                    try:
                        text = element.text.strip()
                        if text:
                            numbers = re.findall(r'\d+', text)
                            if numbers:
                                product_data['price'] = ''.join(numbers)
                                break
                    except:
                        continue
                
                if not product_data['price']:
                    price_elements = self.driver.find_elements(By.XPATH, "//*[contains(., '₽')]")
                    for element in price_elements:
                        try:
                            text = element.text.strip()
                            if text and any(c.isdigit() for c in text):
                                numbers = re.findall(r'\d+', text)
                                if numbers:
                                    product_data['price'] = ''.join(numbers)
                                    break
                        except:
                            continue
            except:
                pass
            
            self.products_data.append(product_data)
            
            status_icons = ["✓" if product_data[key] else "✗" for key in ['sku', 'name', 'price']]
            print(f"  Результат: SKU{status_icons[0]} Назв{status_icons[1]} Цена{status_icons[2]}")
            
            if product_data['name']:
                name_preview = product_data['name'][:60] + "..." if len(product_data['name']) > 60 else product_data['name']
                print(f"  {name_preview}")
            
            self.human_like_pause(1.0, 2.0)
            
        except Exception as e:
            print(f"  Ошибка: {e}")
            self.products_data.append({'sku': '', 'name': '', 'price': '', 'seller_url': self.current_seller_url})
    
    def create_output_folder(self):
        """Создание папки для сохранения файлов"""
        if not os.path.exists(self.output_folder):
            os.makedirs(self.output_folder)
    
    def save_to_excel(self):
        """Сохранение данных в Excel файл"""
        if not self.products_data:
            print("Нет данных для сохранения")
            return
        
        self.create_output_folder()
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"ozon_products_{timestamp}.xlsx"
        filepath = os.path.join(self.output_folder, filename)
        
        df = pd.DataFrame(self.products_data, columns=['sku', 'name', 'price', 'seller_url'])
        df = df.rename(columns={
            'sku': 'SKU',
            'name': 'Название',
            'price': 'Цена с соинвестом',
            'seller_url': 'URL продавца'
        })
        
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Товары')
                worksheet = writer.sheets['Товары']
                worksheet.column_dimensions['A'].width = 20
                worksheet.column_dimensions['B'].width = 70
                worksheet.column_dimensions['C'].width = 20
                worksheet.column_dimensions['D'].width = 40
            
            print(f"\n" + "="*60)
            print("РЕЗУЛЬТАТЫ СОХРАНЕНЫ")
            print("="*60)
            print(f"Файл: {filepath}")
            print(f"Всего товаров: {len(df)}")
            
            if len(df) > 0:
                sku_filled = df['SKU'].notna().sum() + df['SKU'].ne('').sum()
                name_filled = df['Название'].notna().sum() + df['Название'].ne('').sum()
                price_filled = df['Цена с соинвестом'].notna().sum() + df['Цена с соинвестом'].ne('').sum()
                
                print(f"\nСТАТИСТИКА:")
                print(f"С артикулом: {sku_filled}/{len(df)} ({sku_filled/len(df)*100:.1f}%)")
                print(f"С названием: {name_filled}/{len(df)} ({name_filled/len(df)*100:.1f}%)")
                print(f"С ценой: {price_filled}/{len(df)} ({price_filled/len(df)*100:.1f}%)")
            
        except Exception as e:
            print(f"Ошибка при сохранении: {e}")
    
    def run(self):
        """Основной метод запуска парсера"""
        try:
            print("="*70)
            print("ПАРСЕР OZON - С ОБЯЗАТЕЛЬНЫМ РЕШЕНИЕМ КАПЧИ")
            print("="*70)
            print(f"Время начала: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"Количество продавцов для обработки: {len(self.seller_urls)}")
            print("="*70)
            
            # Настройка драйвера
            print("\nНастройка драйвера...")
            self.setup_driver()
            
            # Обрабатываем каждого продавца
            for seller_idx, seller_url in enumerate(self.seller_urls, 1):
                print(f"\n" + "="*70)
                print(f"ПРОДАВЕЦ {seller_idx}/{len(self.seller_urls)}")
                print(f"URL: {seller_url}")
                print("="*70)
                
                self.current_seller_url = seller_url
                seller_name = seller_url.split('/')[-2] if seller_url.endswith('/') else seller_url.split('/')[-1]
                
                # Открываем страницу продавца с обработкой капчи
                print(f"\nОткрываем страницу продавца...")
                if not self.safe_get(seller_url):
                    print(f"⚠ Не удалось загрузить страницу продавца: {seller_name}")
                    continue
                
                # Принимаем куки (только для первого продавца)
                if seller_idx == 1:
                    try:
                        cookie_selectors = [
                            "//button[contains(., 'Принять')]",
                            "//button[contains(., 'Согласен')]",
                        ]
                        
                        for selector in cookie_selectors:
                            try:
                                buttons = self.driver.find_elements(By.XPATH, selector)
                                for button in buttons:
                                    if button.is_displayed():
                                        button.click()
                                        print("✓ Приняты куки")
                                        self.human_like_pause(1, 2)
                                        break
                            except:
                                continue
                    except:
                        pass
                
                # Загружаем товары
                product_links = self.load_all_products_humanlike(seller_name)
                
                if not product_links:
                    print(f"\n⚠ Не найдено товаров у продавца: {seller_name}")
                    continue
                
                print(f"\n" + "="*70)
                print(f"НАЧИНАЕМ ПАРСИНГ {len(product_links)} ТОВАРОВ")
                print("="*70)
                
                # Парсим каждый товар
                start_time = time.time()
                total_to_parse = len(product_links)
                
                for i, link in enumerate(product_links, 1):
                    print(f"\n[{i:3d}/{total_to_parse}] ", end="")
                    self.parse_product_page(link)
                    
                    if i % 10 == 0:
                        elapsed = time.time() - start_time
                        items_per_minute = i / (elapsed / 60)
                        print(f"\n  Прогресс: {i}/{total_to_parse} ({i/total_to_parse*100:.1f}%)")
                        print(f"  Скорость: {items_per_minute:.1f} товаров/мин")
                
                print(f"\n✓ Продавец {seller_name} обработан")
                print(f"✓ Товаров обработано: {total_to_parse}")
            
            print(f"\n" + "="*70)
            print("ВСЕ ПРОДАВЦЫ ОБРАБОТАНЫ!")
            print(f"Время окончания: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            print("="*70)
            
            # Сохраняем результаты
            self.save_to_excel()
            
        except Exception as e:
            print(f"\nКРИТИЧЕСКАЯ ОШИБКА: {e}")
            import traceback
            traceback.print_exc()
        finally:
            if self.driver:
                self.driver.quit()
                print("\n✓ Браузер закрыт")


def create_example_config():
    """Создание примера конфигурационного файла"""
    config_content = """# Файл со списком продавцов Ozon для парсинга
# Каждый URL должен быть на отдельной строке
# Пустые строки и строки начинающиеся с # игнорируются

# Пример URL продавца:
https://www.ozon.ru/seller/energy-strong-172995/

# Можно добавить несколько продавцов:
# https://www.ozon.ru/seller/example-seller-123456/
# https://www.ozon.ru/seller/another-seller-789012/
"""
    
    config_file = 'sellers_list.txt'
    if not os.path.exists(config_file):
        with open(config_file, 'w', encoding='utf-8') as f:
            f.write(config_content)
        print(f"Создан пример конфигурационного файла: {config_file}")
        print("Добавьте в него URL продавцов и запустите парсер снова.")
        return True
    return False


def main():
    """Основная функция запуска парсера"""
    print("="*70)
    print("ПАРСЕР OZON - АВТОМАТИЧЕСКИЙ ЗАПУСК С КАПЧЕЙ")
    print("="*70)
    
    # Проверяем наличие конфигурационного файла
    config_file = 'sellers_list.txt'
    
    if not os.path.exists(config_file):
        print(f"Конфигурационный файл '{config_file}' не найден.")
        create_example_config()
        return
    
    # Читаем файл
    print(f"Чтение конфигурационного файла: {config_file}")
    
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        urls = []
        for line in lines:
            line = line.strip()
            if line and not line.startswith('#'):
                urls.append(line)
        
        if not urls:
            print("⚠ В файле не найдено корректных URL")
            print("Добавьте URL продавцов в формате:")
            print("https://www.ozon.ru/seller/energy-strong-172995/")
            return
        
        print(f"Найдено URL для обработки: {len(urls)}")
        for url in urls:
            print(f"  - {url}")
        
        print("\n" + "="*70)
        print("ВНИМАНИЕ: При обнаружении капчи парсер ОСТАНОВИТСЯ")
        print("и будет ждать, пока вы решите капчу вручную в браузере.")
        print("После решения капчи нажмите Enter в консоли.")
        print("="*70)
        
        confirm = input("\nНачать парсинг? (y/n): ").strip().lower()
        
        if confirm != 'y':
            print("Парсинг отменен.")
            return
        
        # Запускаем парсер
        parser = OzonSellerParser(config_file, output_folder='prices_with_co-investment')
        parser.run()
        
    except Exception as e:
        print(f"Ошибка при запуске парсера: {e}")
        import traceback
        traceback.print_exc()


# Запуск парсера
if __name__ == "__main__":
    main()