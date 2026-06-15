
[app]
title = Automacao Pedagio KM
package.name = pedagio
package.domain = br.app.pedagio
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,xlsx
version = 1.0.0
requirements = python3,kivy==2.3.0,pillow,openpyxl,pdf2image,certifi
orientation = portrait
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.minapi = 21
android.ndk = 25b
android.sdk = 33
android.accept_sdk_license = True
android.arch = arm64-v8a
fullscreen = 0
log_level = 2

[buildozer]
log_level = 2
warn_on_root = 1
