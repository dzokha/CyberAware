# Office
## 1. Key Management Service (KMS) Activator Office
1.1. Thực hiện CMD trong đường dẫn sau với quyền Administrator:

- C:\Program Files (x86)\Microsoft Office\Office14\
- C:\Program Files\Microsoft Office\Office14\

1.2. Kiểm tra trạng thái bản quyền Office
```
cscript //nologo ospp.vbs /dstatus
```

1.3. Gỡ 5 ký tự cuối của các product key đã cài
```
cscript //nologo ospp.vbs /unpkey:B9HB6
```

1.4. Lệnh CMD làm sạch dấu vết KMS trái phép
```
cscript //nologo ospp.vbs /remhst
```

1.5. Kích hoạt Server bên ngoài

**Lưu ý:** 
- Key sưu tầm trên Internet: 2KKDC-67TT9-4XT2F-2MG99-B9HB6
- Server kms8.MSGuides.com chưa kiểm chứng, có nguy cơ đính kèm Malware
```
cscript //nologo ospp.vbs /inpkey:2KKDC-67TT9-4XT2F-2MG99-B9HB6
cscript //nologo ospp.vbs /sethst:kms8.MSGuides.com
cscript //nologo ospp.vbs /act
```
