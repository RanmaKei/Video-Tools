# 🎞️ Video Upscale Pipeline (Upscayl + FFmpeg)

A **GPU-aware, resumable, lossless-first video upscaling pipeline** built around  
**Upscayl (Real-ESRGAN)** and **FFmpeg**.

Designed for **long-running jobs**, **multi-GPU systems**, and **failure-safe recovery** 

---

## ✨ Features

- ✅ Batch processing of videos from a source folder
- ✅ GPU selection (manual or automatic)
- ✅ GPU-tagged work/output folders (`gpu0`, `gpu1`) to avoid collisions
- ✅ Smart resume
  - Keeps extracted frames
  - Only upscales missing frames if interrupted
- ✅ Dry-run mode to inspect incomplete jobs
- ✅ Output validation before cleanup
- ✅ Optional temp retention (`-NoDelete`)
- ✅ Multiple output presets (lossless masters or delivery formats)
- ✅ Safe overwrite control (`-Force`)
- ✅ Audio automatically restored from source
- ✅ **Size-first delivery defaults** (AV1 / HEVC) while keeping **lossless masters** available

---

## 📁 Folder Structure

```
.
├── models
├── source/
│   └── video1.webm
├── _ffv1_work_gpu0/
│   └── video1/
│       ├── frames/
│       ├── upscaled/
│       └── _todo_missing/
├── ffv1_1080p_gpu0/
│   └── video1_ffv1_1080p_gpu0.mkv
├── Video-Upscale.ps1
└── README.md
```

---

## 🔧 Requirements

- Windows  
- PowerShell **7+ recommended**  
- `ffmpeg` and `ffprobe` available in `PATH`  
- Upscayl CLI (`upscayl-bin.exe`)  
- Vulkan-capable GPU (AMD / NVIDIA recommended)

---

## 🚀 Quick Start

```powershell
.\Video-Upscale.ps1
```

### Recommended defaults (best size / modern delivery)
If you want **best file size by default**, set your script default preset to **AV1**:

```powershell
# in param(...)
[string]$OutputPreset = "av1_1080p_mkv"
```

If you need more compatibility (older TVs / devices), default to **HEVC** instead:

```powershell
[string]$OutputPreset = "hevc_1080p_mp4"
```

### Lossless master workflow (recommended archival)
Run a lossless master first, then create delivery encodes later:

```powershell
.\Video-Upscale.ps1 -OutputPreset ffv1_1080p_mkv
```

---

## 🎛️ Output Presets

### Best File Size (Recommended)
| Preset | Description |
|------|------------|
| `av1_1080p_mkv` | AV1 + Opus (MKV) |
| `av1_1080p_mp4` | AV1 + AAC (MP4) |
| `av1_4k_mkv` | 4K AV1 + Opus |
| `av1_4k_mp4` | 4K AV1 + AAC |

### Standard Delivery
| Preset | Description |
|------|------------|
| `hevc_1080p_mp4` | HEVC + AAC |
| `h264_1080p_mp4` | H.264 + AAC |
| `hevc_4k_mp4` | 4K HEVC |
| `h264_4k_mp4` | 4K H.264 |

### Archival / Editing
| Preset | Description |
|------|------------|
| `prores_1080p_mov` | ProRes 422 HQ |
| `prores_4k_mov` | 4K ProRes |
| `ffv1_1080p_mkv` | Lossless FFV1 |
| `ffv1_4k_mkv` | 4K FFV1 |


---

## 🔁 Smart Resume

- Detects incomplete jobs automatically
- Only processes missing frames
- Safe to resume after crash or reboot

---

## 🧪 Validation & Safety

- Output verified via `ffprobe`
- Cleanup only after validation succeeds
- Temp files preserved on failure

---

## 🔗 References

- FFmpeg: https://ffmpeg.org/  
- Upscayl: https://github.com/upscayl/upscayl  
- Real-ESRGAN: https://github.com/xinntao/Real-ESRGAN  
