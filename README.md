# Video-Upscale.ps1

A **GPU-aware video upscaling pipeline** built on **Upscayl + FFmpeg**, designed for long-running jobs, interrupted resumes, and archival-quality outputs.

---

## ✨ Features

### Core Pipeline
1. **Extract frames** from source video (FFmpeg)
2. **Upscale frames** with Upscayl (Real-ESRGAN family)
3. **Normalize geometry** only if needed (optional FFmpeg step)
4. **Rebuild final video** with chosen codec/container
5. **Validate output**
6. **Clean up safely** (or keep temps)

---

## 🔧 Requirements

- Windows  
- PowerShell **7+ recommended**  
- `ffmpeg` and `ffprobe` available in `PATH`  
- Upscayl CLI (`upscayl-bin.exe`)  
- Vulkan-capable GPU (AMD / NVIDIA recommended)

---

## 🔁 Smart Resume

- Detects incomplete jobs automatically
- Only processes missing frames
- Safe to resume after crash or reboot

---

## 📋 Requirements

### External Tools
- `ffmpeg`
- `ffprobe`
- `upscayl-bin.exe`

If not installed system-wide, enable:
```powershell
-AutoDownloadTools
```

---

## 🚀 Usage

### Basic Run
```powershell
pwsh .\Video-Upscale.ps1
```

---

### Resume an Interrupted Job
```powershell
pwsh .\Video-Upscale.ps1 -Resume
```

---

### Dry-Run (No Work Performed)
```powershell
pwsh .\Video-Upscale.ps1 -DryRun
```

---

### Keep Temporary Files
```powershell
pwsh .\Video-Upscale.ps1 -NoDelete
```

---

## 📁 Directory Layout

```
Video-Tools/
├─ Video-Upscale.ps1
├─ source/
├─ tools/
├─ _ffv1_work_gpu0/
└─ output_av1_1080p_gpu0/
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

## 🧪 Validation & Safety

- Output verified via `ffprobe`
- Cleanup only after validation succeeds
- Temp files preserved on failure

---

## 🔗 References

- FFmpeg: https://ffmpeg.org/  
- Upscayl: https://github.com/upscayl/upscayl  
- Real-ESRGAN: https://github.com/xinntao/Real-ESRGAN  

