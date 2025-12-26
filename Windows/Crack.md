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
# Windows
1. Hiển thị thông tin cơ bản: edition, kênh bản quyền (Retail / MAK / KMS)
```
slmgr /dli
```
2. Hiển thị chi tiết đầy đủ: KMS server, Thời hạn kích hoạt, License status
```
slmgr /dlv
```
3. Kiểm tra: Windows đã kích hoạt vĩnh viễn hay chưa, Hay chỉ tạm thời (KMS)
```
slmgr /xpr
```
4. Xóa cấu hình KMS server
```
slmgr /ckms
```
5. Reset thời gian đánh giá (grace period) của Windows
```
slmgr /rearm
```
6. Xóa product key khỏi Windows Registry
```
slmgr /cpky
```
7. Gỡ product key khỏi hệ thống.
```
slmgr /upk
```
8. Yêu cầu Windows kích hoạt theo cấu hình

**Lưu ý:** 
- Key sưu tầm trên Internet: W269N-WFGWX-YVC9B-4J6C9-T83GX
- Server kms8.MSGuides.com chưa kiểm chứng, có nguy cơ đính kèm Malware
```
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX
slmgr /skms kms8.MSGuides.com
slmgr /ato
```
9. Cài license từ file .xrm-ms
```
slmgr /ilc <license_file>
```
