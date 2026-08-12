import sys
import re
import os

def update_version():
    raw_ver = sys.argv[1] if len(sys.argv) > 1 else 'v2.1.4'
    ver = raw_ver.lstrip('v') if raw_ver and raw_ver.startswith('v') else raw_ver
    if not ver or ver == 'main' or ver == 'master':
        ver = '2.1.4'

    config_path = os.path.join('Daftar_Radiologi', 'core', 'config.py')
    with open(config_path, 'r', encoding='utf-8') as f:
        content = f.read()

    new_content = re.sub(r'APP_VERSION\s*=\s*".*?"', f'APP_VERSION = "{ver}"', content)

    with open(config_path, 'w', encoding='utf-8') as f:
        f.write(new_content)

    print(f'[AutoVersion] Injected APP_VERSION = "{ver}" into {config_path}')

if __name__ == '__main__':
    update_version()
