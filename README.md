# 🌐 IP Location Checker

> **[VI]** Kiểm tra vị trí địa lý hàng loạt IP — không cần Python, không cần cài thêm gì — xuất kết quả ra file Excel đẹp.
>
> **[EN]** Bulk IP geolocation lookup — no Python, no dependencies — exports clean Excel results in minutes.

---

## 📋 Giới thiệu / Introduction

**[VI]** **IP Location Checker** là bộ công cụ chạy trực tiếp trên Windows, cho phép tra cứu thông tin địa lý (quốc gia, vùng, thành phố, ISP...) của hàng nghìn địa chỉ IP cùng lúc, hoàn toàn tự động.

**[EN]** **IP Location Checker** is a lightweight Windows tool that performs bulk geolocation lookups (country, region, city, ISP, ASN...) on thousands of IP addresses simultaneously — fully automated, no installation required.

| | **Tiếng Việt** | **English** |
|---|---|---|
| Ngôn ngữ / Language | PowerShell (có sẵn trên Windows) | PowerShell (built-in on Windows) |
| API | [ip-api.com](http://ip-api.com) — miễn phí, không cần đăng ký | [ip-api.com](http://ip-api.com) — free, no registration |
| Đầu vào / Input | File `.txt` — mỗi dòng 1 IP | `.txt` file — one IP per line |
| Đầu ra / Output | File `.xlsx` gồm 3 sheet | `.xlsx` file with 3 sheets |
| Hiệu suất / Performance | ~3.000–6.000 IP trong 3–6 phút | ~3,000–6,000 IPs in 3–6 minutes |

---

## 📁 Cấu trúc thư mục / Folder Structure

```
📁 Check IPLocation\
    ├── check_ip_location.cmd   ← File chạy chính / Main launcher (double-click)
    ├── ip_checker.ps1          ← Script PowerShell xử lý logic / Core processing script
    ├── ip_list.txt             ← Danh sách IP cần kiểm tra / IP list to check (user-created)
    └── IP_Location_YYYYMMDD_HHmmss.xlsx  ← Kết quả / Output (auto-generated)
```

---

## ⚙️ Yêu cầu hệ thống / System Requirements

| | **Tiếng Việt** | **English** |
|---|---|---|
| OS | Windows 7 / 10 / 11 | Windows 7 / 10 / 11 |
| Runtime | PowerShell 3.0+ (có sẵn mặc định) | PowerShell 3.0+ (built-in) |
| Mạng / Network | Kết nối Internet | Internet connection |
| Phụ thuộc / Dependencies | ❌ Không cần Python, Excel, v.v. | ❌ No Python, Excel, or extras needed |

---

## 🚀 Hướng dẫn sử dụng / How to Use

### Bước 1 / Step 1 — Chuẩn bị file IP / Prepare the IP file

**[VI]** Tạo file `ip_list.txt` trong cùng thư mục với file `.cmd`, mỗi dòng một địa chỉ IP:

**[EN]** Create `ip_list.txt` in the same folder as the `.cmd` file, one IP address per line:

```
8.8.8.8
1.1.1.1
103.21.244.0
14.191.33.165
...
```

> **[VI]** Lưu ý:
> - Dòng trống hoặc dòng bắt đầu bằng `#` sẽ bị bỏ qua
> - IP trùng lặp sẽ tự động được loại bỏ
> - Hỗ trợ IPv4
>
> **[EN]** Notes:
> - Empty lines and lines starting with `#` are ignored
> - Duplicate IPs are automatically deduplicated
> - IPv4 supported

---

### Bước 2 / Step 2 — Chạy tool / Run the tool

**[VI]** Double-click vào file **`check_ip_location.cmd`**

**[EN]** Double-click **`check_ip_location.cmd`**

```
✅ [VI] Không cần "Run as Administrator"
✅ [VI] Không cần mở PowerShell thủ công
✅ [VI] Nếu chưa có ip_list.txt, tool sẽ hỏi tạo file mẫu 10 IP

✅ [EN] No "Run as Administrator" needed
✅ [EN] No need to open PowerShell manually
✅ [EN] If ip_list.txt is missing, the tool offers to create a 10-IP sample file
```

---

### Bước 3 / Step 3 — Theo dõi tiến trình / Monitor progress

**[VI]** Tool hiển thị progress bar theo thời gian thực:

**[EN]** The tool displays a real-time progress bar:

```
  [############################--------] 72%  Batch 26/35  |  2600/3487 IP  |  ETA: 01:12
```

---

### Bước 4 / Step 4 — Nhận kết quả / Get the result

**[VI]** Sau khi hoàn thành, file Excel được tạo tự động tại cùng thư mục. Tool sẽ hỏi có muốn mở file ngay không.

**[EN]** Once finished, the Excel file is created automatically in the same folder. The tool will ask if you want to open it immediately.

```
IP_Location_20260615_104556.xlsx
```

---

## 📊 Cấu trúc file Excel đầu ra / Output Excel Structure

**[VI]** File Excel gồm **3 sheet** / **[EN]** The Excel file contains **3 sheets**:

---

### Sheet 1 — Summary (Tổng quan / Overview)

| Trường / Field | Mô tả (VI) | Description (EN) |
|---|---|---|
| Ngay gio chay | Thời điểm thực thi | Execution timestamp |
| File IP dau vao | Đường dẫn file input | Path to the input file |
| Tong so IP | Số IP đọc từ file | Total IPs read from file |
| IP thanh cong | Số IP tra cứu thành công | Successfully resolved IPs |
| IP loi | Số IP không xác định được | Failed / unresolved IPs |
| So quoc gia | Số quốc gia phát hiện | Number of countries detected |
| Thoi gian xu ly | Tổng thời gian chạy | Total processing time |

---

### Sheet 2 — IP Location Detail (Chi tiết từng IP / Per-IP Detail)

| Cột / Column | Mô tả (VI) | Description (EN) |
|---|---|---|
| STT | Số thứ tự | Row number |
| IP Address | Địa chỉ IP | IP address |
| Country | Tên quốc gia | Country name |
| Code | Mã quốc gia | Country code (VN, US, SG...) |
| Region | Tỉnh / Vùng | Region / Province |
| City | Thành phố | City |
| ISP | Nhà cung cấp dịch vụ Internet | Internet Service Provider |
| Organization | Tổ chức sở hữu | Owning organization |
| AS Number | Mã hệ thống tự trị | Autonomous System Number |
| Status | Trạng thái tra cứu | Lookup status |

> **[VI]** Hàng màu **đỏ nhạt** = IP lỗi. Hàng **xanh nhạt / trắng** xen kẽ để dễ đọc.
>
> **[EN]** **Light red** rows = failed IPs. **Light blue / white** rows alternate for readability.

---

### Sheet 3 — Country Summary (Thống kê quốc gia / Country Breakdown)

**[VI]** Xếp hạng các quốc gia theo số lượng IP, kèm phần trăm.

**[EN]** Countries ranked by IP count with percentage of total.

---

## ⏱️ Hiệu suất tham khảo / Performance Reference

| Số lượng IP / IP Count | Thời gian (VI) | Est. Time (EN) |
|---|---|---|
| 500 IP | ~1 phút | ~1 minute |
| 1.000 IP | ~1,5 phút | ~1.5 minutes |
| 3.000 IP | ~3 phút | ~3 minutes |
| 6.000 IP | ~5–6 phút | ~5–6 minutes |

> **[VI]** API ip-api.com cho phép tối đa **45 request/phút** với batch 100 IP/request. Tool tự động điều tiết tốc độ để tránh rate limit.
>
> **[EN]** ip-api.com allows up to **45 requests/minute**, with 100 IPs per batch request. The tool automatically throttles to stay within rate limits.

---

## ❓ Xử lý lỗi / Troubleshooting

| Lỗi / Error | Nguyên nhân (VI) | Cause (EN) | Cách xử lý / Fix |
|---|---|---|---|
| `Khong tim thay ip_list.txt` | File input chưa tồn tại | Input file not found | Tạo `ip_list.txt` cùng thư mục / Create `ip_list.txt` in the same folder |
| `Khong tim thay ip_checker.ps1` | Thiếu file script | Script file missing | Đảm bảo cả 2 file `.cmd` và `.ps1` cùng thư mục / Keep both `.cmd` and `.ps1` in the same folder |
| `Request failed` trong cột Status | Mất kết nối hoặc IP không hợp lệ | Network issue or invalid IP | Kiểm tra mạng; IP lỗi vẫn ghi vào Excel / Check network; failed IPs are still recorded |
| PowerShell bị chặn | Group Policy tổ chức | Execution policy blocked | Chạy lệnh / Run: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser` |

---

## 🔒 Bảo mật / Security

**[VI]**
- Tool **không gửi** dữ liệu nào ngoài địa chỉ IP đến API ip-api.com
- API ip-api.com là dịch vụ **miễn phí, công khai**, không yêu cầu đăng nhập
- Không có dữ liệu nào được lưu trên máy chủ bên ngoài

**[EN]**
- The tool **only sends IP addresses** to the ip-api.com API — no other data
- ip-api.com is a **free, public** service — no account or login required
- No data is stored on any external server beyond the API lookup itself

---

## 📞 Thông tin API / API Reference

Tool sử dụng / Uses: **[ip-api.com Batch API](http://ip-api.com/docs/api:batch)**

| | **Tiếng Việt** | **English** |
|---|---|---|
| Endpoint | `http://ip-api.com/batch` | `http://ip-api.com/batch` |
| Giới hạn / Limit | 45 requests/phút, 100 IP/request | 45 requests/min, 100 IPs/request |
| Chi phí / Cost | Miễn phí | Free |
| Xác thực / Auth | Không cần API key | No API key required |

---

*Phiên bản / Version: 1.0 — PowerShell thuần, không phụ thuộc thư viện ngoài / Pure PowerShell, zero external dependencies.*
