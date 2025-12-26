# Office
Key Management Service (KMS) Activator Office

1. Thực hiện CMD trong đường dẫn sau với quyền Administrator:

- C:\Program Files (x86)\Microsoft Office\Office14\
- C:\Program Files\Microsoft Office\Office14\

2. Các lệnh cơ bản
```
cscript //nologo ospp.vbs /dstatus        REM Kiểm tra trạng thái bản quyền Office
cscript //nologo ospp.vbs /remhst         REM Lệnh CMD làm sạch dấu vết KMS trái phép
cscript //nologo ospp.vbs /unpkey:B9HB6   REM Gỡ 5 ký tự cuối của các product key đã cài
```
5. Kích hoạt Server bên ngoài

**Lưu ý:** 
- Key sưu tầm trên Internet: 2KKDC-67TT9-4XT2F-2MG99-B9HB6
- Server kms8.MSGuides.com chưa kiểm chứng, có nguy cơ đính kèm Malware
```
cscript //nologo ospp.vbs /inpkey:2KKDC-67TT9-4XT2F-2MG99-B9HB6
cscript //nologo ospp.vbs /sethst:kms8.MSGuides.com
cscript //nologo ospp.vbs /act
```
# Windows
1. Các câu lệnh cơ bản
```
slmgr /dli     REM Hiển thị thông tin cơ bản
slmgr /dlv     REM Hiển thị chi tiết đầy đủ
slmgr /xpr     REM Kiểm tra: Windows đã kích hoạt
slmgr /ckms    REM Xóa cấu hình KMS server
slmgr /cpky    REM Xóa product key khỏi Windows Registry
slmgr /upk     REM Gỡ product key khỏi hệ thống
slmgr /rearm   REM Reset thời gian đánh giá (grace period) của Windows
```
2. Yêu cầu Windows kích hoạt theo cấu hình

**Lưu ý:** 
- Key sưu tầm trên Internet: W269N-WFGWX-YVC9B-4J6C9-T83GX
- Server kms8.MSGuides.com chưa kiểm chứng, có nguy cơ đính kèm Malware
```
slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX    REM Cài product key
slmgr /skms kms8.MSGuides.com               REM Trỏ đến KMS server
slmgr /ato                                  REM Yêu cầu Windows kích hoạt
```
3. Cài license từ file
```
slmgr /ilc <license_file>   REM file .xrm-ms
```
