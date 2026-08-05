import os
import shutil
import json
import datetime
import sqlite3


def create_backup(db_path='forest_data.db', backup_dir='backups'):
    os.makedirs(backup_dir, exist_ok=True)
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    backup_name = f'forestapp_backup_{timestamp}'
    backup_path = os.path.join(backup_dir, backup_name)
    os.makedirs(backup_path, exist_ok=True)

    if os.path.exists(db_path):
        shutil.copy2(db_path, os.path.join(backup_path, 'forest_data.db'))

    metadata = {
        'app': 'ForestApp',
        'version': '2.0',
        'created_at': timestamp,
        'tables': [],
    }
    try:
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        for row in cursor.fetchall():
            table = row[0]
            cursor.execute(f'SELECT COUNT(*) FROM "{table}"')
            count = cursor.fetchone()[0]
            metadata['tables'].append({'name': table, 'rows': count})
        conn.close()
    except Exception:
        pass

    with open(os.path.join(backup_path, 'metadata.json'), 'w', encoding='utf-8') as f:
        json.dump(metadata, f, ensure_ascii=False, indent=2)

    archive_path = os.path.join(backup_dir, f'{backup_name}.zip')
    shutil.make_archive(backup_path, 'zip', backup_path)
    shutil.rmtree(backup_path)
    return archive_path


def restore_backup(zip_path, db_path='forest_data.db'):
    import tempfile
    import zipfile

    with zipfile.ZipFile(zip_path, 'r') as zf:
        backup_db = None
        for name in zf.namelist():
            if name.endswith('forest_data.db'):
                backup_db = name
                break
        if not backup_db:
            raise FileNotFoundError('В архиве не найден forest_data.db')

        with tempfile.TemporaryDirectory() as tmpdir:
            zf.extract(backup_db, tmpdir)
            src = os.path.join(tmpdir, backup_db)
            if os.path.exists(db_path):
                os.replace(db_path, db_path + '.bak')
            shutil.copy2(src, db_path)
    return True


def list_backups(backup_dir='backups'):
    if not os.path.exists(backup_dir):
        return []
    backups = []
    for f in sorted(os.listdir(backup_dir), reverse=True):
        if f.endswith('.zip') and f.startswith('forestapp_backup_'):
            full_path = os.path.join(backup_dir, f)
            size = os.path.getsize(full_path)
            meta_path = None
            meta_name = f.replace('.zip', '') + '.json'
            alt_dir = os.path.join(backup_dir, f.replace('.zip', ''))
            if os.path.exists(os.path.join(backup_dir, meta_name)):
                meta_path = os.path.join(backup_dir, meta_name)
            elif os.path.exists(os.path.join(alt_dir, 'metadata.json')):
                meta_path = os.path.join(alt_dir, 'metadata.json')

            meta = {}
            if meta_path:
                try:
                    with open(meta_path, 'r', encoding='utf-8') as mf:
                        meta = json.load(mf)
                except Exception:
                    pass

            backups.append({
                'path': full_path,
                'name': f,
                'size': size,
                'created': f.replace('forestapp_backup_', '').replace('.zip', ''),
                'metadata': meta,
            })
    return backups
