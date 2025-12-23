# OpenChrome (Magic.BrowserAutomationNET)

OpenChrome adalah helper library untuk memastikan instance Chrome + Selenium
berada dalam kondisi usable, dengan ownership dan lifecycle yang jelas.

Library ini application-agnostic dan tidak bergantung pada:
- UI
- AppState
- Domain aplikasi tertentu

Dapat digunakan di:
- WinForms
- Console
- Service / Worker
- Background automation

---

## Tujuan Utama

Library ini dibuat untuk menyelesaikan masalah klasik Selenium:

- Chrome terlihat masih ada, tapi driver sudah mati
- Chrome zombie dengan user-data-dir terkunci
- Tidak jelas siapa yang bertanggung jawab menutup Chrome
- Risiko menutup Chrome yang masih dipakai task lain

OpenChrome menyelesaikan masalah tersebut dengan state ownership yang eksplisit.

---

## Konsep Inti: Ownership Chrome

OpenChrome membedakan dua kondisi awal Chrome:

| State | Arti |
|------|------|
| Opened | Chrome sudah ada & bisa dipakai |
| NotOpened | Chrome tidak ada / zombie, dibuka oleh OpenChrome |

Ownership ini disimpan di properti:

StateCode ChromeInitialState

Prinsip:
- Yang membuka Chrome → bertanggung jawab menutup
- Yang hanya meminjam → tidak boleh menutup

---

## Fitur

- Mengecek apakah Chrome masih terkoneksi Selenium
- Membunuh Chrome zombie berdasarkan UserDataDir
- Membuka Chrome baru jika diperlukan
- Event lifecycle untuk observability
- Tidak mengatur registry, pool, atau threading

---

## Public API

### Constructor

OpenChrome(Chrome chrome)

Parameter:
- chrome      : Instance Chrome yang dikelola caller

---

### Start()

Chrome Start()

Fungsi:
- Mengecek koneksi Chrome
- Menghubungkan ke Chrome yang masih hidup
- Membuka Chrome baru jika diperlukan

Return:
- **Chrome** → Instance Chrome yang **siap digunakan**
  - Bisa berupa Chrome lama (reuse berhasil)
  - Atau Chrome baru (dibuka oleh OpenChrome)

Catatan:
- Method ini **tidak pernah mengembalikan null**
- Semua kegagalan koneksi ditangani internal


---

### ForceClose()

void ForceClose()

Menutup Chrome tanpa memperhatikan ownership.

Gunakan untuk:
- Shutdown
- Emergency cleanup

---

### DisposeIfOwned()

void DisposeIfOwned()

Menutup Chrome hanya jika Chrome dibuka oleh instance OpenChrome ini
(yaitu saat ChromeInitialState == NotOpened).

Ini adalah cara paling aman untuk lifecycle normal.

---

## Nilai yang Dikembalikan Library

Library ini **tidak mengembalikan status kompleks lewat return value**, melainkan melalui:
- Return object (`Chrome`)
- Properti state
- Event lifecycle

---

### 1. Return Langsung

#### `Chrome Start()`

- Mengembalikan instance `Chrome` yang valid dan usable
- Caller **tidak perlu membedakan** Chrome lama atau baru

---

### 2. Properti State

#### `StateCode ChromeInitialState`

Menunjukkan **kondisi Chrome saat Start() pertama kali dipanggil**:

| Nilai | Arti | Dampak Lifecycle |
|------|------|------------------|
| Opened | Chrome sudah ada & valid | Caller **tidak boleh** menutup Chrome |
| NotOpened | Chrome tidak ada / zombie | Caller **wajib** menutup via DisposeIfOwned() |

Properti ini menjadi **sumber kebenaran ownership Chrome**.

---

### 3. Event Lifecycle (Callback)

Library menyediakan event untuk observability:

OpenChromeEvents

EventType:
- Start
- ChromeIsConnected
- ZombieKilled
- NewChromeIsOpened
- BrowserClosed

Semua event membawa payload:
- `Chrome` → instance Chrome aktif

---

## Event Lifecycle

Library menyediakan event untuk observability:

OpenChromeEvents

EventType:
- Start
- ChromeIsConnected
- ZombieKilled
- NewChromeIsOpened
- BrowserClosed

---

## Contoh Penggunaan

```csharp
var openChrome = new OpenChrome(chrome, userDataDir);

openChrome.OpenChromeEvents += e =>
{
    Console.WriteLine($"Event: {e.EventType}");
};

var usableChrome = openChrome.Start();

// lakukan automation di sini

openChrome.DisposeIfOwned();
```

---

## Dispose Pattern

OpenChrome:
- sealed
- Tidak memiliki unmanaged resource
- Tidak menggunakan finalizer

Method Dispose() hanya menandai lifecycle object,
penutupan Chrome dilakukan secara eksplisit melalui:
- ForceClose()
- DisposeIfOwned()

---

## Design Principles

- Single Responsibility
- Explicit Ownership
- Application-agnostic
- Deterministic lifecycle
- Automation-safe

---

## Catatan Penting

- Library ini tidak thread-safe secara internal
- Sinkronisasi paralel Chrome adalah tanggung jawab caller
- Jangan share satu instance Chrome ke banyak thread

---

## Lisensi

Internal / Private library
Gunakan sesuai kebutuhan proyek
