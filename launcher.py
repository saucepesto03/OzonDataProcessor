import sys
import os

def clear():
    os.system("cls" if os.name == "nt" else "clear")

def pause():
    input("\nНажмите Enter для возврата в меню...")

def run_parser():
    clear()
    print("ЗАПУСК: Парсер Ozon\n")
    try:
        from parser import main
        main()
    except Exception as e:
        print(f"Ошибка запуска парсера: {e}")
    pause()

def run_dimensions():
    clear()
    print("ЗАПУСК: Загрузка габаритов Ozon API\n")
    try:
        from dimensions import main
        main()
    except Exception as e:
        print(f"Ошибка запуска dimensions: {e}")
    pause()

def run_unit_update():
    clear()
    print("ЗАПУСК: Обновление Unit-файла\n")
    try:
        from script import update_unit_file
        success = update_unit_file()
        if success:
            print("\n✓ Обновление успешно")
        else:
            print("\n✗ Обновление завершилось с ошибками")
    except Exception as e:
        print(f"Ошибка запуска unit_update: {e}")
    pause()

# ========== ДОБАВЛЕНА НОВАЯ ФУНКЦИЯ ==========
def run_full_pipeline():
    clear()
    print("=" * 60)
    print("🚀 ЗАПУСК ПОЛНОЙ ЦЕПОЧКИ")
    print("API → Парсинг → Обновление Unit")
    print("=" * 60)
    
    try:
        # ШАГ 1: Загрузка габаритов
        print("\n" + "=" * 60)
        print("ШАГ 1/3: Загрузка габаритов (Ozon API)")
        print("=" * 60)
        from dimensions import main as dimensions_main
        dimensions_main()
        
        # ШАГ 2: Парсинг продавцов
        print("\n" + "=" * 60)
        print("ШАГ 2/3: Парсинг продавцов Ozon")
        print("=" * 60)
        from parser import main as parser_main
        parser_main()
        
        # ШАГ 3: Обновление Unit
        print("\n" + "=" * 60)
        print("ШАГ 3/3: Обновление Unit-файла")
        print("=" * 60)
        from script import update_unit_file
        success = update_unit_file()
        
        print("\n" + "=" * 60)
        if success:
            print("✅ ВСЯ ЦЕПОЧКА ВЫПОЛНЕНА УСПЕШНО!")
        else:
            print("❌ Ошибка на этапе обновления Unit")
        print("=" * 60)

    except KeyboardInterrupt:
        print("\n\n⚠️  Цепочка прервана пользователем")
    except Exception as e:
        print(f"\n❌ КРИТИЧЕСКАЯ ОШИБКА ЦЕПОЧКИ: {e}")
        import traceback
        traceback.print_exc()
    
    pause()
# =============================================

def main_menu():
    while True:
        clear()
        print("=" * 60)
        print("OZON TOOLKIT — ЕДИНЫЙ ЛАУНЧЕР")
        print("=" * 60)
        print("1. Парсер продавцов Ozon (Selenium)")
        print("2. Получить габариты товаров (Ozon API)")
        print("3. Обновить Unit-файл (Excel)")
        print("4. 🚀 Полная цепочка (API → Парсинг → Unit)")
        print("0. Выход")
        print("=" * 60)

        choice = input("Выберите пункт: ").strip()

        if choice == "1":
            run_parser()
        elif choice == "2":
            run_dimensions()
        elif choice == "3":
            run_unit_update()
        elif choice == "4":  # НОВЫЙ ПУНКТ
            run_full_pipeline()
        elif choice == "0":
            print("\nВыход.")
            sys.exit(0)
        else:
            print("\nНеверный выбор")
            pause()

if __name__ == "__main__":
    main_menu()