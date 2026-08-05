[app]
title = Фанаты Пихты
package.name = forestapp
package.domain = org.forestcrypt
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,ttf,json,docx,xlsx
version = 1.0.0
requirements = python3,kivy==2.3.1,kivymd==1.2.0,pandas,openpyxl,python-docx,Pillow,sqlite3,pyjnius,android
orientation = portrait
fullscreen = 0
android.permissions = WRITE_EXTERNAL_STORAGE,READ_EXTERNAL_STORAGE,INTERNET
android.api = 31
android.minapi = 21
android.ndk = 25b
android.ndk_api = 21
android.private_storage = True
android.logcat_filters = *:S python:D
android.copy_libs = 1
android.archs = arm64-v8a,armeabi-v7a
android.allow_backup = True

[buildozer]
profile = android
log_level = 2
warn_on_root = 1
