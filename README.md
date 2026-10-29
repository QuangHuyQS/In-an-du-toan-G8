# Tool in an du toan G8 (V6)

Xuat PDF tu Excel du toan theo FILE TONG HOP: di theo link trong file tong toi tung file chi tiet,
tach theo hang muc, xuat TH chi phi TB / THKP / Gia tong hop / (Don gia chi tiet) / Gia thang.

- `src\` : ma nguon VBA (UTF-8). `inanDTG8_V6.bas` la engine, API chinh `API_RunTong`.
- `BUILD_V6.ps1` : sinh `gen\` roi nap vao file .xlsm (can Excel + Trust access to VBA project).
  Lenh: `.\BUILD_V6.ps1 -KeepForm`
- `FILE_TONG_IN_G8_V6.4.xlsm` : ban build moi nhat.
