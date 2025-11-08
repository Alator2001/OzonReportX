"""
Автоматическое обновление программы с GitHub
"""
import requests
import sys
import os
import zipfile
import shutil
import subprocess
from pathlib import Path
from packaging import version

# Конфигурация
GITHUB_REPO = "Alator2001/OzonReportX"
CURRENT_VERSION = "1.0.0"
VERSION_FILE = Path(__file__).resolve().parent.parent / "version.txt"


def get_current_version():
    """Получает текущую версию из файла или константы"""
    if VERSION_FILE.exists():
        return VERSION_FILE.read_text(encoding='utf-8').strip()
    return CURRENT_VERSION


def get_latest_release():
    """Получает информацию о последнем релизе с GitHub"""
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(api_url, timeout=10)
        
        if response.status_code == 404:
            print("Релизы не найдены на GitHub")
            return None
            
        response.raise_for_status()
        data = response.json()
        
        return {
            'tag': data['tag_name'].lstrip('v'),
            'download_url': data['zipball_url'],
            'name': data.get('name', 'Без названия'),
            'notes': data.get('body', 'Нет описания')
        }
    except requests.exceptions.RequestException as e:
        print(f"Не удалось подключиться к GitHub: {e}")
        return None
    except Exception as e:
        print(f"Ошибка при получении информации о релизе: {e}")
        return None


def download_and_extract(download_url, temp_dir):
    """Скачивает и распаковывает релиз"""
    print("Скачивание обновления...")
    
    try:
        response = requests.get(download_url, timeout=60, stream=True)
        response.raise_for_status()
        
        zip_path = temp_dir / "update.zip"
        
        # Скачиваем с прогресс-баром
        total_size = int(response.headers.get('content-length', 0))
        downloaded = 0
        
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if total_size > 0:
                        percent = (downloaded / total_size) * 100
                        print(f"\rПрогресс: {percent:.1f}%", end='')
        
        print("\n✓ Скачивание завершено")
        
        # Распаковываем
        print("Распаковка...")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        
        # Находим распакованную папку (GitHub создаёт папку с названием repo-hash)
        extracted_dirs = [d for d in temp_dir.iterdir() if d.is_dir() and d.name != '__MACOSX']
        if not extracted_dirs:
            raise RuntimeError("Не найдена распакованная папка")
        
        return extracted_dirs[0]
        
    except Exception as e:
        raise RuntimeError(f"Ошибка при скачивании: {e}")


def create_backup(repo_root):
    """Создаёт резервную копию текущей версии"""
    backup_dir = repo_root.parent / f"backup_v{get_current_version()}"
    
    print(f"Создание резервной копии в {backup_dir.name}...")
    
    if backup_dir.exists():
        shutil.rmtree(backup_dir)
    
    # Копируем всё, кроме виртуального окружения и временных файлов
    shutil.copytree(
        repo_root, 
        backup_dir,
        ignore=shutil.ignore_patterns(
            '.venv', '__pycache__', '*.pyc', '.git', 
            'reports', '.env', '*.log', 'backup_*'
        )
    )
    
    print(f"✓ Резервная копия создана: {backup_dir}")
    return backup_dir


def apply_update(source_dir, repo_root):
    """Применяет обновление, копируя файлы"""
    print("Установка обновления...")
    
    # Список файлов и папок, которые НЕ нужно обновлять
    preserve = {'.venv', '.git', '__pycache__', 'reports', '.env', 
                'costs.xlsx', 'costs.csv', 'version.txt', 'updater.log',
                'backup_*'}
    
    updated_count = 0
    
    for item in source_dir.iterdir():
        # Пропускаем защищённые файлы
        if any(item.match(pattern) for pattern in preserve):
            continue
        
        dest = repo_root / item.name
        
        try:
            if dest.exists():
                if dest.is_dir():
                    shutil.rmtree(dest)
                else:
                    dest.unlink()
            
            if item.is_dir():
                shutil.copytree(item, dest)
            else:
                shutil.copy2(item, dest)
            
            updated_count += 1
            
        except Exception as e:
            print(f"⚠ Не удалось обновить {item.name}: {e}")
    
    print(f"✓ Обновлено файлов: {updated_count}")


def update_version_file(new_version):
    """Обновляет файл с номером версии"""
    VERSION_FILE.write_text(new_version, encoding='utf-8')


def check_and_update():
    """Основная функция проверки и обновления"""
    print("Проверка обновлений...")
    
    current = get_current_version()
    print(f"Текущая версия: {current}")
    
    # Получаем информацию о последнем релизе
    latest_info = get_latest_release()
    
    if not latest_info:
        print("✓ Не удалось проверить обновления")
        return False
    
    latest = latest_info['tag']
    print(f"Последняя версия: {latest}")
    
    # Сравниваем версии
    try:
        if version.parse(latest) <= version.parse(current):
            print("✓ Используется последняя версия")
            return False
    except Exception as e:
        print(f"Ошибка при сравнении версий: {e}")
        return False
    
    # Есть обновление
    print(f"\n🎉 Доступно обновление: {latest_info['name']}")
    print(f"\nЧто нового:\n{latest_info['notes']}\n")
    
    # Спрашиваем пользователя
    answer = input("Установить обновление сейчас? [Y/n]: ").strip().lower()
    if answer and answer not in ('y', 'yes', 'д', 'да'):
        print("Обновление пропущено")
        return False
    
    # Процесс обновления
    repo_root = Path(__file__).resolve().parent.parent
    temp_dir = Path.home() / "AppData" / "Local" / "Temp" / "ozonreportx_update"
    
    try:
        # Создаём временную папку
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
        temp_dir.mkdir(parents=True, exist_ok=True)
        
        # Создаём резервную копию
        backup_dir = create_backup(repo_root)
        
        # Скачиваем и распаковываем
        source_dir = download_and_extract(latest_info['download_url'], temp_dir)
        
        # Применяем обновление
        apply_update(source_dir, repo_root)
        
        # Обновляем версию
        update_version_file(latest)
        
        # Очищаем временные файлы
        shutil.rmtree(temp_dir)
        
        print(f"\n✓ Обновление до версии {latest} завершено успешно!")
        print(f"Резервная копия: {backup_dir}")
        print("\nПрограмма будет перезапущена...\n")
        
        # Перезапускаем программу
        os.execv(sys.executable, [sys.executable] + sys.argv)
        
    except Exception as e:
        print(f"\n❌ Ошибка при обновлении: {e}")
        print(f"Вы можете восстановить программу из резервной копии")
        return False
    
    return True


if __name__ == "__main__":
    try:
        check_and_update()
    except KeyboardInterrupt:
        print("\n\nОбновление отменено пользователем")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        sys.exit(1)
