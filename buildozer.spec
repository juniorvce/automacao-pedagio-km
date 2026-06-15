[app]
title = Automacao Pedagio KM
package.name = pedagio
package.domain = br.app.pedagio
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 1.0.0
requirements = python3,kivy==2.3.1,pillow,openpyxl
orientation = portrait
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 34
android.minapi = 21
android.ndk = 25b
android.accept_sdk_license = True
android.archs = arm64-v8a
fullscreen = 0

# Trava o python-for-android numa release estavel (usa Python 3.11, compativel com Kivy)
p4a.branch = 2024.01.21
p4a.fork = kivy

[buildozer]
log_level = 2
warn_on_root = 1
