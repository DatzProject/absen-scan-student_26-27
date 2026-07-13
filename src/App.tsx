import React, { useState, useEffect, useRef } from "react";
import jsQR from "jsqr";
// @ts-ignore
import QRCode from "qrcode";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
// @ts-ignore
import { Canvg } from "canvg";
import SignatureCanvas from "react-signature-canvas";
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend,
  ChartData,
  ChartOptions,
  ChartEvent,
  LegendItem,
} from "chart.js";
import { Bar } from "react-chartjs-2";

ChartJS.register(
  CategoryScale,
  LinearScale,
  BarElement,
  Title,
  Tooltip,
  Legend
);

const endpoint =
  "https://script.google.com/macros/s/AKfycbyPlDeLjd2W4SjYkuVSGZylOa_jpRd1UUmwetsu_Fn8EYS9YdkpLK-tnkWuE6cUWcpg/exec";
const SHEET_SEMESTER1 = "RekapSemester1";
const SHEET_SEMESTER2 = "RekapSemester2";

interface Student {
  id: string;
  name: string;
  nisn: string; // ← Ubah dari string | null | undefined menjadi string
  kelas: string;
  jenisKelamin: string;
  foto?: string;
}

interface SchoolData {
  namaKepsek: string;
  nipKepsek: string;
  ttdKepsek: string;
  namaGuru: string;
  nipGuru: string;
  ttdGuru: string;
  namaKota: string;
  statusGuru: string;
  namaSekolah: string;
}

type AttendanceStatus = "Hadir" | "Izin" | "Sakit" | "Alpha";

interface AttendanceRecord {
  [date: string]: {
    [studentId: string]: AttendanceStatus | "";
  };
}

interface KeteranganRecord {
  [date: string]: {
    [studentId: string]: string;
  };
}

interface MonthlyRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

interface GraphData {
  [month: string]: {
    Hadir: number;
    Alpha: number;
    Izin: number;
    Sakit: number;
  };
}

interface StatusSummary {
  Hadir: number;
  Izin: number;
  Sakit: number;
  Alpha: number;
}

interface StatusVisibility {
  Hadir: boolean;
  Alpha: boolean;
  Izin: boolean;
  Sakit: boolean;
}

interface AttendanceHistory {
  tanggal: string;
  nama: string;
  kelas: string;
  nisn: string;
  status: AttendanceStatus;
}

interface SemesterRecap {
  nama: string;
  kelas: string;
  hadir: number;
  alpa: number;
  izin: number;
  sakit: number;
  persenHadir: number;
}

interface TanggalMerah {
  tanggal: string;
  deskripsi: string;
  tanggalAkhir?: string;
}

interface JadwalMengajar {
  kelas: string;
  hari: string;
}

const formatDateDDMMYYYY = (isoDate: string): string => {
  const [year, month, day] = isoDate.split("-");
  return `${day}/${month}/${year}`;
};

type EditedRecord = {
  date: string;
  nisn: string;
  status: AttendanceStatus | "";
};

const AttendanceTab: React.FC<{
  students: Student[];
  onRecapRefresh: () => void;
  onLoadingChange: (loading: boolean) => void;
  studentsLoaded: boolean;
}> = ({ students, onRecapRefresh, onLoadingChange, studentsLoaded }) => {
  const [attendance, setAttendance] = useState<AttendanceRecord>({});

  const getLocalDate = () => {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  };

  const [date, setDate] = useState<string>(getLocalDate());
  const [selectedKelas, setSelectedKelas] = useState<string>("");
  const [showDebugInfo, setShowDebugInfo] = useState<boolean>(false);
  const [isSaving, setIsSaving] = useState<boolean>(false);

  // QR Scanner states
  const [scanMode, setScanMode] = useState<boolean>(false);
  const [scannedStudents, setScannedStudents] = useState<Set<string>>(
    new Set()
  );
  const videoRef = useRef<HTMLVideoElement>(null);
  const [scannerError, setScannerError] = useState<string>("");
  const [lastScannedId, setLastScannedId] = useState<string>("");
  const [scannedStudentPhotos, setScannedStudentPhotos] = useState<
    Array<{ student: Student; timestamp: Date }>
  >([]);
  const scannerIntervalRef = useRef<number | null>(null);
  const sendingLockRef = useRef<Set<string>>(new Set());
  const [isSendingData, setIsSendingData] = useState<boolean>(false);

  const [tanggalMerahList, setTanggalMerahList] = useState<TanggalMerah[]>([]);
  const [loadingTanggalMerah, setLoadingTanggalMerah] =
    useState<boolean>(false);
  const [jadwalMengajar, setJadwalMengajar] = useState<JadwalMengajar[]>([]);
  const [loadingJadwal, setLoadingJadwal] = useState<boolean>(false);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [existingStudentIds, setExistingStudentIds] = useState<Set<string>>(
    new Set()
  );
  const [isLoadingExistingData, setIsLoadingExistingData] =
    useState<boolean>(false);
  const [existingAttendanceData, setExistingAttendanceData] = useState<any[]>(
    []
  );
  const [fotoAbsensiUrl, setFotoAbsensiUrl] = useState<string | null>(null);
  const [isUploadingFoto, setIsUploadingFoto] = useState<boolean>(false);
  const [isLoadingFoto, setIsLoadingFoto] = useState<boolean>(false);
  const [previewFoto, setPreviewFoto] = useState<string | null>(null);
  const [fotoToUpload, setFotoToUpload] = useState<string | null>(null);
  const [isDeletingFoto, setIsDeletingFoto] = useState<boolean>(false);
  const [showFotoFullscreen, setShowFotoFullscreen] = useState<boolean>(false);
  const [keterangan, setKeterangan] = useState<{
    [date: string]: { [studentId: string]: string };
  }>({});

  const getFotoUrlAttendance = (fotoUrl: string | undefined): string => {
    if (!fotoUrl) return "";
    if (fotoUrl.includes("lh3.googleusercontent.com")) return fotoUrl;
    if (fotoUrl.includes("uc?export=view&id=")) {
      const fileId = new URLSearchParams(fotoUrl.split("?")[1]).get("id");
      if (fileId) return `https://lh3.googleusercontent.com/d/${fileId}`;
    }
    return fotoUrl;
  };

  const uniqueClasses = React.useMemo(() => {
    const classSet = new Set<string>();
    students.forEach((student) => {
      let kelasValue = student.kelas;
      if (kelasValue != null) {
        kelasValue = String(kelasValue).trim();
        if (
          kelasValue !== "" &&
          kelasValue !== "undefined" &&
          kelasValue !== "null"
        ) {
          classSet.add(kelasValue);
        }
      }
    });
    const classes = Array.from(classSet).sort((a, b) => {
      const aIsNum = /^\d+$/.test(a);
      const bIsNum = /^\d+$/.test(b);
      if (aIsNum && bIsNum) return parseInt(a) - parseInt(b);
      if (aIsNum && !bIsNum) return -1;
      if (!aIsNum && bIsNum) return 1;
      return a.localeCompare(b);
    });
    return [...classes];
  }, [students]);

  useEffect(() => {
    if (uniqueClasses.length > 0 && selectedKelas === "") {
      setSelectedKelas(uniqueClasses[0]);
    }
  }, [uniqueClasses]);

  const filteredStudents = React.useMemo(() => {
    if (!selectedKelas) return students;
    return students.filter((student) => {
      if (student.kelas == null) return false;
      const studentKelas = String(student.kelas).trim();
      return studentKelas === selectedKelas;
    });
  }, [students, selectedKelas]);

  // ===== QR SCANNER FUNCTIONS =====

  const startScanner = async () => {
    try {
      setScannerError("");
      setScanMode(true);

      const stream = await navigator.mediaDevices.getUserMedia({
        video: { facingMode: "environment" },
      });

      if (videoRef.current) {
        videoRef.current.srcObject = stream;
        await videoRef.current.play();

        scannerIntervalRef.current = window.setInterval(() => {
          if (!scannerIntervalRef.current) return;
          scanQRCode();
        }, 500);
      }
    } catch (error) {
      console.error("Scanner error:", error);
      setScannerError(
        "Tidak dapat mengakses kamera. Pastikan izin kamera telah diberikan."
      );
      stopScanner();
    }
  };

  const scanQRCode = async () => {
    if (!videoRef.current) return;

    try {
      const canvas = document.createElement("canvas");
      const video = videoRef.current;
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      const ctx = canvas.getContext("2d");
      if (!ctx) return;

      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
      const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);

      const code = jsQR(imageData.data, imageData.width, imageData.height);

      if (code && code.data) {
        handleScanSuccess(code.data);
      }
    } catch (error) {
      console.error("QR scan error:", error);
    }
  };

  const stopScanner = () => {
    if (scannerIntervalRef.current) {
      clearInterval(scannerIntervalRef.current);
      scannerIntervalRef.current = null;
    }

    if (videoRef.current && videoRef.current.srcObject) {
      const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
      tracks.forEach((track) => track.stop());
      videoRef.current.srcObject = null;
    }

    setScanMode(false);

    // ✅ TAMBAHKAN: Hapus popup dan overlay saat stop scanner
    const popup = document.getElementById("scan-popup");
    const overlay = document.getElementById("scan-overlay");

    if (popup) popup.remove();
    if (overlay) overlay.remove();

    console.log("🛑 Scanner stopped dan popup/overlay dibersihkan");
  };

  const getCurrentTime = () => {
    const now = new Date();
    return `${String(now.getHours()).padStart(2, "0")}:${String(
      now.getMinutes()
    ).padStart(2, "0")}`;
  };

  // ===== FUNGSI KIRIM DATA LANGSUNG =====
  const sendAttendanceData = async (student: Student) => {
    if (sendingLockRef.current.has(student.nisn)) {
      console.log(`🔒 ${student.name} sudah dikunci, skip pengiriman`);
      return;
    }

    sendingLockRef.current.add(student.nisn);
    setIsSendingData(true);

    const formattedDate = formatDateDDMMYYYY(date);

    // ✅ PERBAIKAN: Pastikan keterangan diambil dengan benar
    const keteranganValue =
      (keterangan[date] && keterangan[date][student.id]) || "";

    const jamSekarang = getCurrentTime(); //

    console.log("📤 Mengirim data:", {
      // TAMBAHKAN LOG
      nama: student.name,
      keterangan: keteranganValue,
    });

    const data = [
      {
        tanggal: formattedDate,
        nama: student.name || "N/A",
        kelas: student.kelas || "N/A",
        nisn: student.nisn || "N/A",
        status: "Hadir",
        keterangan: keteranganValue,
        jam: getCurrentTime(),
      },
    ];

    try {
      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });

      console.log(
        `✅ Data ${student.name} berhasil dikirim dengan keterangan:`,
        keteranganValue
      );
      loadExistingAttendanceData(false);
      onRecapRefresh();
    } catch (error) {
      console.error("❌ Gagal kirim data:", error);
      alert(`❌ Gagal menyimpan data ${student.name}. Silakan coba lagi.`);
      sendingLockRef.current.delete(student.nisn);
      setScannedStudents((prev) => {
        const newSet = new Set(prev);
        newSet.delete(student.id);
        return newSet;
      });
    } finally {
      setIsSendingData(false);
    }
  };

  const handleScanSuccess = async (qrData: string) => {
    try {
      // Prevent duplicate scans
      if (qrData === lastScannedId) return;
      setLastScannedId(qrData);
      setTimeout(() => setLastScannedId(""), 2000);

      // Extract NISN from QR data (format: NISN:1234567890)
      let nisn = qrData;
      if (qrData.includes("NISN:")) {
        nisn = qrData.split("NISN:")[1];
      }

      // Find student by NISN
      const student = filteredStudents.find(
        (s) => String(s.nisn).trim() === nisn.trim()
      );

      if (!student) {
        playErrorSound();
        alert(`❌ Siswa dengan NISN ${nisn} tidak ditemukan!`);
        return; // interval tetap jalan, scanner lanjut
      }

      // Check if already scanned
      if (scannedStudents.has(student.id)) {
        playErrorSound();
        alert(`⚠️ ${student.name} sudah diabsen sebelumnya!`);
        return; // interval tetap jalan, scanner lanjut
      }

      // Check if already has data in database
      if (existingStudentIds.has(student.id)) {
        playErrorSound();
        alert(`⚠️ ${student.name} sudah memiliki data absensi di database!`);
        return; // interval tetap jalan, scanner lanjut
      }

      // Hentikan interval HANYA jika scan berhasil (siswa valid & belum diabsen)
      if (scannerIntervalRef.current) {
        clearInterval(scannerIntervalRef.current);
        scannerIntervalRef.current = null;
      }

      // Set attendance locally
      setAttendance((prev) => ({
        ...prev,
        [date]: {
          ...prev[date],
          [student.id]: "Hadir",
        },
      }));

      // Mark as scanned
      setScannedStudents((prev) => new Set([...prev, student.id]));

      // Tambahkan ke list foto (cegah duplikat)
      setScannedStudentPhotos((prev) => {
        const sudahAda = prev.some((item) => item.student.id === student.id);
        if (sudahAda) return prev;
        return [{ student, timestamp: new Date() }, ...prev];
      });

      // Tutup kamera
      if (videoRef.current && videoRef.current.srcObject) {
        const tracks = (videoRef.current.srcObject as MediaStream).getTracks();
        tracks.forEach((track) => track.stop());
        videoRef.current.srcObject = null;
      }
      setScanMode(false);

      // Success feedback
      playSuccessSound();
      showSuccessNotification(student.name, student.foto);

      // Kirim data ke server
      await sendAttendanceData(student);
    } catch (error) {
      console.error("Scan processing error:", error);
      playErrorSound();
      alert("❌ Error memproses QR Code");
    }
  };

  const handleUploadFotoAbsensi = async () => {
    try {
      // Buat input file yang hanya menerima foto dari kamera
      const input = document.createElement("input");
      input.type = "file";
      input.accept = "image/*";
      input.capture = "environment"; // Langsung buka kamera

      input.onchange = async (e: any) => {
        const file = e.target.files?.[0];
        if (!file) {
          return;
        }

        try {
          // Kompres foto
          const compressedBase64 = await compressImage(file, 2); // Max 2MB

          // Tampilkan preview
          setPreviewFoto(compressedBase64);
          setFotoToUpload(compressedBase64.split(",")[1]); // Simpan base64 tanpa header
        } catch (error) {
          console.error("Error processing foto:", error);
          alert("❌ Gagal memproses foto");
        }
      };

      input.click();
    } catch (error) {
      console.error("Error uploading foto:", error);
      alert("❌ Gagal mengupload foto");
    }
  };

  // Fungsi untuk konfirmasi upload
  const confirmUploadFoto = async () => {
    if (!fotoToUpload) return;

    try {
      setIsUploadingFoto(true);

      const formattedDate = formatDateDDMMYYYY(date);
      const uploadData = {
        type: "uploadFotoAbsensi",
        tanggal: formattedDate,
        kelas: selectedKelas,
        fotoBase64: fotoToUpload,
        mimeType: "image/jpeg",
      };

      const response = await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(uploadData),
      });

      console.log("✅ Foto absensi berhasil diupload");
      alert(
        `✅ Foto absensi kelas ${selectedKelas} tanggal ${formattedDate} berhasil disimpan!`
      );

      // Reset preview dan reload foto
      setPreviewFoto(null);
      setFotoToUpload(null);
      loadFotoAbsensi();
    } catch (error) {
      console.error("Error uploading foto:", error);
      alert("❌ Gagal mengupload foto");
    } finally {
      setIsUploadingFoto(false);
    }
  };

  // Fungsi untuk cancel upload
  const cancelUploadFoto = () => {
    setPreviewFoto(null);
    setFotoToUpload(null);
  };

  // Fungsi kompresi gambar
  const compressImage = (file: File, maxSizeMB: number): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();

      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement("canvas");
          let width = img.width;
          let height = img.height;

          // Hitung ukuran maksimal (misal max 1920px untuk sisi terpanjang)
          const maxDimension = 1920;
          if (width > height && width > maxDimension) {
            height = (height * maxDimension) / width;
            width = maxDimension;
          } else if (height > maxDimension) {
            width = (width * maxDimension) / height;
            height = maxDimension;
          }

          canvas.width = width;
          canvas.height = height;

          const ctx = canvas.getContext("2d");
          if (!ctx) {
            reject(new Error("Canvas context not available"));
            return;
          }

          ctx.drawImage(img, 0, 0, width, height);

          // Kompresi bertahap sampai ukuran di bawah maxSizeMB
          let quality = 0.9;
          let compressedDataUrl = canvas.toDataURL("image/jpeg", quality);

          // Hitung ukuran dalam MB
          const getFileSizeMB = (dataUrl: string) => {
            const base64 = dataUrl.split(",")[1];
            const bytes = atob(base64).length;
            return bytes / (1024 * 1024);
          };

          // Kurangi quality sampai ukuran memenuhi syarat
          while (
            getFileSizeMB(compressedDataUrl) > maxSizeMB &&
            quality > 0.1
          ) {
            quality -= 0.1;
            compressedDataUrl = canvas.toDataURL("image/jpeg", quality);
          }

          console.log(
            `Foto dikompres ke ${getFileSizeMB(compressedDataUrl).toFixed(
              2
            )} MB (quality: ${quality.toFixed(1)})`
          );
          resolve(compressedDataUrl);
        };

        img.onerror = () => reject(new Error("Failed to load image"));
        img.src = e.target?.result as string;
      };

      reader.onerror = () => reject(new Error("Failed to read file"));
      reader.readAsDataURL(file);
    });
  };

  const loadFotoAbsensi = async () => {
    setIsLoadingFoto(true);
    setFotoAbsensiUrl(null); // ← reset dulu sebelum fetch
    try {
      const formattedDate = formatDateDDMMYYYY(date);
      const url = `${endpoint}?action=fotoAbsensi&tanggal=${formattedDate}&kelas=${selectedKelas}`;
      const response = await fetch(url, { method: "GET", mode: "cors" });

      if (response.ok) {
        const result = await response.json();
        if (result.success) {
          setFotoAbsensiUrl(result.fotoUrl);
        } else {
          setFotoAbsensiUrl(null);
        }
      } else {
        setFotoAbsensiUrl(null);
      }
    } catch (error) {
      console.error("Error loading foto absensi:", error);
      setFotoAbsensiUrl(null);
    } finally {
      setIsLoadingFoto(false); // ← selesai loading
    }
  };

  const handleDeleteFoto = async () => {
    if (
      !confirm(
        `⚠️ Yakin ingin menghapus foto absensi kelas ${selectedKelas} tanggal ${formatDateDDMMYYYY(
          date
        )}?\n\nTindakan ini tidak dapat dibatalkan!`
      )
    ) {
      return;
    }

    setIsDeletingFoto(true);

    try {
      const formattedDate = formatDateDDMMYYYY(date);
      const deleteData = {
        type: "deleteFotoAbsensi",
        tanggal: formattedDate,
        kelas: selectedKelas,
      };

      await fetch(endpoint, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(deleteData),
      });

      console.log("✅ Foto absensi berhasil dihapus");
      alert(
        `✅ Foto absensi kelas ${selectedKelas} tanggal ${formattedDate} berhasil dihapus!`
      );

      // Reset state dan reload
      setFotoAbsensiUrl(null);
      loadFotoAbsensi();
    } catch (error) {
      console.error("Error deleting foto:", error);
      alert("❌ Gagal menghapus foto");
    } finally {
      setIsDeletingFoto(false);
    }
  };

  const playSuccessSound = () => {
    const audio = new Audio(
      "data:audio/wav;base64,UklGRnoGAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQoGAACBhYqFbF1fdJivrJBhNjVgodDbq2EcBj+a2/LDciUFLIHO8tiJNwgZaLvt559NEAxQp+PwtmMcBjiR1/LMeSwFJHfH8N2QQAoUXrTp66hVFApGn+DyvmwhBSuBzvLZiTcIGWi77eefTRAMUKfj8LZjHAY4ktfyy3ksBSR3x/DdkEAKFF606+uoVRQKRp/g8r5sIQUrgc7y2Yk3CBlou+3nn00QDFCn4/C2YxwGOJLX8st5LAUkd8fw3ZBACh=="
    );
    audio.play().catch(() => {});
  };

  const playErrorSound = () => {
    const audio = new Audio(
      "data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAESsAACJWAAABAAgAZGF0YQAAAAA="
    );
    audio.play().catch(() => {});
  };

  const showSuccessNotification = (name: string, fotoUrl?: string) => {
    // ✅ PERBAIKAN: Hapus popup dan overlay lama jika masih ada
    const existingPopup = document.getElementById("scan-popup");
    const existingOverlay = document.getElementById("scan-overlay");

    if (existingPopup) {
      existingPopup.remove();
      console.log("🗑️ Popup lama dihapus");
    }

    if (existingOverlay) {
      existingOverlay.remove();
      console.log("🗑️ Overlay lama dihapus");
    }

    const resolvedFoto = fotoUrl
      ? fotoUrl.includes("lh3.googleusercontent.com")
        ? fotoUrl
        : fotoUrl.includes("uc?export=view&id=")
        ? `https://lh3.googleusercontent.com/d/${new URLSearchParams(
            fotoUrl.split("?")[1]
          ).get("id")}`
        : fotoUrl
      : "";

    // ✅ TAMBAHKAN: Buat overlay DULU sebelum popup
    const overlay = document.createElement("div");
    overlay.id = "scan-overlay";
    overlay.style.cssText = `
      position: fixed;
      inset: 0;
      background: rgba(0,0,0,0.5);
      z-index: 9998;
    `;

    // ✅ PERBAIKAN: Fungsi untuk cleanup yang aman
    const cleanup = () => {
      const popup = document.getElementById("scan-popup");
      const overlay = document.getElementById("scan-overlay");

      if (popup) {
        popup.style.opacity = "0";
        popup.style.transition = "opacity 0.2s";
        setTimeout(() => {
          if (popup.parentNode) {
            popup.remove();
          }
        }, 200);
      }

      if (overlay) {
        overlay.style.opacity = "0";
        overlay.style.transition = "opacity 0.2s";
        setTimeout(() => {
          if (overlay.parentNode) {
            overlay.remove();
          }
        }, 200);
      }

      console.log("✅ Cleanup popup dan overlay selesai");
    };

    overlay.addEventListener("click", cleanup);

    const popup = document.createElement("div");
    popup.id = "scan-popup";
    popup.style.cssText = `
      position: fixed;
      top: 50%;
      left: 50%;
      transform: translate(-50%, -50%);
      background: white;
      border-radius: 16px;
      box-shadow: 0 20px 60px rgba(0,0,0,0.4);
      z-index: 9999;
      padding: 24px;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 12px;
      border: 4px solid #4ade80;
      min-width: 320px;
      animation: popupFadeIn 0.3s ease;
    `;

    // Tambahkan animasi
    const style = document.createElement("style");
    style.textContent = `
      @keyframes popupFadeIn {
        from { opacity: 0; transform: translate(-50%, -50%) scale(0.8); }
        to { opacity: 1; transform: translate(-50%, -50%) scale(1); }
      }
    `;

    // ✅ PERBAIKAN: Hapus style lama jika ada
    const existingStyle = document.getElementById("popup-animation-style");
    if (existingStyle) {
      existingStyle.remove();
    }
    style.id = "popup-animation-style";
    document.head.appendChild(style);

    popup.innerHTML = `
      <div style="
        background: #4ade80;
        color: white;
        padding: 6px 20px;
        border-radius: 20px;
        font-weight: bold;
        font-size: 16px;
      ">✅ HADIR</div>
  
      ${
        resolvedFoto
          ? `<img
            src="${resolvedFoto}"
            style="
              width: 340px;
              height: 255px;
              object-fit: cover;
              border-radius: 12px;
              border: 3px solid #4ade80;
              box-shadow: 0 4px 20px rgba(0,0,0,0.15);
            "
            onerror="this.style.display='none'; this.nextElementSibling.style.display='flex';"
          />
          <div style="
            display: none;
            width: 340px;
            height: 255px;
            background: #f0fdf4;
            border-radius: 12px;
            border: 3px solid #4ade80;
            align-items: center;
            justify-content: center;
            font-size: 80px;
          ">👤</div>`
          : `<div style="
            display: flex;
            width: 340px;
            height: 255px;
            background: #f0fdf4;
            border-radius: 12px;
            border: 3px solid #4ade80;
            align-items: center;
            justify-content: center;
            font-size: 80px;
          ">👤</div>`
      }
  
      <div style="
        font-size: 22px;
        font-weight: bold;
        color: #1f2937;
        text-align: center;
        max-width: 340px;
      ">${name}</div>
  
      <div style="
        font-size: 13px;
        color: #6b7280;
      ">Tap untuk menutup</div>
    `;

    // ✅ PERBAIKAN: Klik untuk tutup dengan cleanup yang aman
    popup.addEventListener("click", cleanup);

    // ✅ PERBAIKAN: Append overlay DULU, baru popup
    document.body.appendChild(overlay);
    document.body.appendChild(popup);

    // ✅ PERBAIKAN: Auto tutup setelah 4 detik dengan cleanup
    const autoCloseTimer = setTimeout(() => {
      cleanup();
    }, 8000);

    // ✅ TAMBAHKAN: Simpan timer ID untuk bisa di-clear jika ditutup manual
    popup.dataset.timerId = String(autoCloseTimer);

    // ✅ PERBAIKAN: Clear timer jika ditutup manual
    popup.addEventListener("click", () => {
      const timerId = popup.dataset.timerId;
      if (timerId) {
        clearTimeout(Number(timerId));
      }
    });

    console.log("✅ Popup dan overlay berhasil dibuat");
  };

  // Cleanup on unmount
  useEffect(() => {
    return () => {
      stopScanner();
    };
  }, []);

  // Load existing attendance data
  const loadExistingAttendanceData = async (showLoading: boolean = true) => {
    setIsLoadingExistingData(true);
    if (showLoading) onLoadingChange(true);
    setExistingStudentIds(new Set());
    setExistingAttendanceData([]);

    try {
      const formattedDate = formatDateDDMMYYYY(date);
      // ✅ Kirim parameter tanggal dan kelas ke server
      const url = `${endpoint}?action=attendanceHistory&tanggal=${formattedDate}&kelas=${selectedKelas}`;
      const controller = new AbortController();
      const timeoutId = setTimeout(() => controller.abort(), 15000);
      const response = await fetch(url, {
        method: "GET",
        mode: "cors",
        signal: controller.signal,
      });
      clearTimeout(timeoutId);

      if (response.ok) {
        const result = await response.json();
        if (result.success && result.data) {
          // ✅ Tidak perlu filter lagi di client karena sudah difilter server
          const filteredData = result.data;

          setExistingAttendanceData(filteredData);

          const existingAttendanceRecord: { [key: string]: AttendanceStatus } =
            {};
          const existingKeteranganRecord: { [key: string]: string } = {};
          const existingIds = new Set<string>();

          filteredStudents.forEach((student) => {
            const existingRecord = filteredData.find(
              (record: any) => record.nama === student.name
            );
            if (existingRecord) {
              existingAttendanceRecord[student.id] =
                existingRecord.status as AttendanceStatus;
              existingKeteranganRecord[student.id] =
                existingRecord.keterangan || "";
              existingIds.add(student.id);
            } else {
              existingAttendanceRecord[student.id] = "" as any;
              existingKeteranganRecord[student.id] = "";
            }
          });

          setAttendance((prev) => ({
            ...prev,
            [date]: existingAttendanceRecord,
          }));
          setKeterangan((prev) => ({
            ...prev,
            [date]: existingKeteranganRecord,
          }));
          setExistingStudentIds(existingIds);
        }
      }
    } catch (error) {
      console.error("Error loading existing attendance data:", error);
    } finally {
      setIsLoadingExistingData(false);
      onLoadingChange(false);
    }
  };

  const allStudentsHaveData =
    filteredStudents.length > 0 &&
    existingStudentIds.size === filteredStudents.length;

  useEffect(() => {
    if (students.length > 0) {
      loadExistingAttendanceData();
      setScannedStudentPhotos([]);
      setScannedStudents(new Set());
      sendingLockRef.current = new Set(); // ← TAMBAHKAN
    } else {
      // Tidak ada siswa sama sekali — matikan loading supaya
      // tidak nyangkut selamanya di "Mohon Tunggu..."
      onLoadingChange(false);
    }
  }, [date, selectedKelas, students]);

  useEffect(() => {
    if (students.length && !attendance[date]) {
      const init: { [key: string]: AttendanceStatus } = {};
      students.forEach((s) => (init[s.id] = "" as any));
      setAttendance((prev) => ({ ...prev, [date]: init }));
    }
  }, [date, students, attendance]);

  useEffect(() => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => (res.ok ? res.json() : Promise.reject()))
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        }
      })
      .catch(console.error);
  }, []);

  useEffect(() => {
    if (students.length > 0 && selectedKelas !== "Semua") {
      loadFotoAbsensi();
    } else {
      setFotoAbsensiUrl(null);
    }
  }, [date, selectedKelas, students, allStudentsHaveData]); // ✅ tambah allStudentsHaveData

  useEffect(() => {
    // Cleanup function saat component unmount
    return () => {
      // Hapus semua popup dan overlay yang mungkin masih tersisa
      const popup = document.getElementById("scan-popup");
      const overlay = document.getElementById("scan-overlay");
      const style = document.getElementById("popup-animation-style");

      if (popup) popup.remove();
      if (overlay) overlay.remove();
      if (style) style.remove();

      console.log("🧹 Cleanup saat component unmount");
    };
  }, []);

  useEffect(() => {
    fetchTanggalMerah();
    fetchJadwalMengajar();
  }, []);

  const fetchTanggalMerah = async () => {
    setLoadingTanggalMerah(true);
    try {
      const res = await fetch(`${endpoint}?action=tanggalMerah`);
      const data = await res.json();
      if (data.success) setTanggalMerahList(data.data || []);
    } catch (error) {
      console.error("Error fetch tanggal merah:", error);
    } finally {
      setLoadingTanggalMerah(false);
    }
  };

  const fetchJadwalMengajar = async () => {
    setLoadingJadwal(true);
    try {
      const res = await fetch(`${endpoint}?action=jadwalMengajar`);
      const data = await res.json();
      if (data.success) setJadwalMengajar(data.data || []);
    } catch (error) {
      console.error("Error fetch jadwal mengajar:", error);
    } finally {
      setLoadingJadwal(false);
    }
  };

  const formatDateDDMMYYYY = (isoDate: string): string => {
    const [year, month, day] = isoDate.split("-");
    return `${day}/${month}/${year}`;
  };

  const setStatus = (sid: string, status: AttendanceStatus) => {
    if (existingStudentIds.has(sid)) return;
    setAttendance((prev) => ({
      ...prev,
      [date]: { ...prev[date], [sid]: status },
    }));
  };

  const setKeteranganValue = (sid: string, value: string) => {
    if (existingStudentIds.has(sid)) return;
    setKeterangan((prev) => {
      const currentDateData = prev[date] || {};
      return {
        ...prev,
        [date]: {
          ...currentDateData,
          [sid]: value,
        },
      };
    });
  };

  const handleSave = () => {
    setIsSaving(true);
    const formattedDate = formatDateDDMMYYYY(date);
    const studentsToSave = filteredStudents.filter(
      (s) => !existingStudentIds.has(s.id)
    );

    if (studentsToSave.length === 0) {
      alert(
        "✅ Semua siswa sudah diabsen. Tidak ada data baru untuk disimpan."
      );
      setIsSaving(false);
      return;
    }

    const jamSekarang = getCurrentTime();

    const data = studentsToSave.map((s) => {
      // ✅ PERBAIKAN: Ambil keterangan dengan benar
      const keteranganValue =
        (keterangan[date] && keterangan[date][s.id]) || "";

      console.log("📤 Data untuk", s.name, "keterangan:", keteranganValue); // LOG

      return {
        tanggal: formattedDate,
        nama: s.name || "N/A",
        kelas: s.kelas || "N/A",
        nisn: s.nisn || "N/A",
        status: attendance[date]?.[s.id] || "Hadir",
        keterangan: keteranganValue,
        jam: attendance[date]?.[s.id] === "Hadir" ? jamSekarang : "",
      };
    });

    console.log("📤 Mengirim data ke server:", data); // LOG UNTUK DEBUG

    fetch(endpoint, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data),
    })
      .then(() => {
        const message = `✅ Data absensi siswa baru kelas ${selectedKelas} berhasil dikirim!`;
        alert(message);
        onRecapRefresh();
        loadExistingAttendanceData();
        setScannedStudents(new Set());
        setIsSaving(false);
      })
      .catch(() => {
        alert("❌ Gagal kirim data absensi.");
        setIsSaving(false);
      });
  };

  const statusColor: Record<AttendanceStatus, string> = {
    Hadir: "bg-green-500",
    Izin: "bg-yellow-400",
    Sakit: "bg-blue-400",
    Alpha: "bg-red-500",
  };

  const getAttendanceSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredStudents.forEach((s) => {
      const status = attendance[date]?.[s.id] as AttendanceStatus;
      summary[status]++;
    });
    return summary;
  };

  const isSunday = (dateStr: string): boolean => {
    const [year, month, day] = dateStr.split("-");
    const date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
    return date.getDay() === 0;
  };

  const isInDateRange = (
    currentDate: string,
    startDate: string,
    endDate?: string
  ): boolean => {
    const formatToComparable = (dateStr: string) => {
      const [d, m, y] = dateStr.split("/");
      return new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
    };
    const [year, month, day] = currentDate.split("-");
    const current = new Date(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day)
    );
    const start = formatToComparable(startDate);
    if (!endDate) return current.getTime() === start.getTime();
    const end = formatToComparable(endDate);
    return current >= start && current <= end;
  };

  const getTanggalMerahInfo = (dateStr: string): TanggalMerah | null => {
    for (const tm of tanggalMerahList) {
      if (isInDateRange(dateStr, tm.tanggal, tm.tanggalAkhir)) return tm;
    }
    return null;
  };

  const isLiburSemester = (dateStr: string): boolean => {
    const info = getTanggalMerahInfo(dateStr);
    if (!info) return false;
    const desc = info.deskripsi.toLowerCase();
    return (
      desc.includes("libur akhir semester") || desc.includes("libur semester")
    );
  };

  const isTanggalMerah = (dateStr: string): boolean => {
    const info = getTanggalMerahInfo(dateStr);
    if (!info) return false;
    const desc = info.deskripsi.toLowerCase();
    return !(
      desc.includes("libur akhir semester") || desc.includes("libur semester")
    );
  };

  const isBukanJadwalMengajar = (dateStr: string): boolean => {
    if (loadingJadwal) return false; // ← TAMBAHKAN
    if (!schoolData) return false;
    if (schoolData?.statusGuru === "Guru Kelas") return false;
    if (selectedKelas === "Semua") return false;
    if (loadingJadwal) return false;
    const jadwal = jadwalMengajar.find((j) => j.kelas === selectedKelas);
    if (!jadwal) return true;
    const [year, month, day] = dateStr.split("-");
    const currentDate = new Date(
      parseInt(year),
      parseInt(month) - 1,
      parseInt(day)
    );
    const dayNames = [
      "Minggu",
      "Senin",
      "Selasa",
      "Rabu",
      "Kamis",
      "Jumat",
      "Sabtu",
    ];
    const dayName = dayNames[currentDate.getDay()];
    const hariJadwal = jadwal.hari
      .split(",")
      .map((h) => h.trim())
      .filter((h) => h.length > 0);
    return !hariJadwal.includes(dayName);
  };

  const attendanceSummary = getAttendanceSummary();

  if (students.length === 0) {
    // Kalau fetch data siswa BELUM selesai → tampilkan loading
    if (!studentsLoaded) {
      return (
        <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
          <div className="bg-white p-6 rounded-lg shadow-md">
            <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
              📋 Absensi Siswa
            </h2>
            <div className="text-center py-12">
              <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-8 max-w-md mx-auto">
                <div className="text-6xl mb-4">⏳</div>
                <h3 className="text-2xl font-bold text-blue-700 mb-2">
                  Mohon Tunggu
                </h3>
                <p className="text-blue-600 mb-4">
                  Sedang memuat data siswa...
                </p>
                <div className="flex justify-center">
                  <svg
                    className="animate-spin h-8 w-8 text-blue-600"
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    ></circle>
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    ></path>
                  </svg>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    // Kalau fetch SUDAH selesai tapi memang tidak ada siswa → tampilkan pesan kosong
    return (
      <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
        <div className="bg-white p-6 rounded-lg shadow-md">
          <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
            📋 Absensi Siswa
          </h2>
          <div className="text-center py-12">
            <div className="bg-gray-50 border-2 border-gray-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">📭</div>
              <h3 className="text-2xl font-bold text-gray-700 mb-2">
                Data Siswa Kosong
              </h3>
              <p className="text-gray-500">
                Belum ada data siswa yang terdaftar. Silakan tambahkan data
                siswa terlebih dahulu.
              </p>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📋 Absensi Siswa dengan QR Code
        </h2>

        {isLoadingExistingData && (
          <div className="mb-4 text-center">
            <p className="text-blue-600 text-sm">⏳ Memuat data absensi...</p>
          </div>
        )}

        {allStudentsHaveData && !isLoadingExistingData && (
          <div className="mb-6 bg-green-50 border border-green-200 rounded-lg p-4 text-center">
            <div className="text-green-700 font-semibold text-lg mb-2">
              ✅ Semua siswa sudah diabsen
            </div>
            <p className="text-green-600 text-sm">
              Data absensi untuk tanggal {formatDateDDMMYYYY(date)}
              {selectedKelas !== "Semua" ? ` kelas ${selectedKelas}` : ""} sudah
              lengkap.
            </p>
          </div>
        )}

        {existingStudentIds.size > 0 &&
          !allStudentsHaveData &&
          !isLoadingExistingData && (
            <div className="mb-6 bg-yellow-50 border border-yellow-200 rounded-lg p-4 text-center">
              <div className="text-yellow-700 font-semibold text-lg mb-2">
                ⚠️ Sebagian siswa sudah diabsen
              </div>
              <p className="text-yellow-600 text-sm">
                {existingStudentIds.size} dari {filteredStudents.length} siswa
                sudah memiliki data absensi.
              </p>
            </div>
          )}

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Tanggal</p>
            <input
              type="date"
              value={date}
              onChange={(e) => setDate(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-gray-100 text-gray-500 cursor-not-allowed"
            />
          </div>

          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>

          {selectedKelas !== "Semua" &&
            !isSunday(date) &&
            !isTanggalMerah(date) &&
            !isLiburSemester(date) && (
              <div className="text-center">
                <p className="text-sm text-gray-500 mb-2">Foto Kelas</p>
                <button
                  onClick={handleUploadFotoAbsensi}
                  disabled={
                    isUploadingFoto || isLoadingFoto || !!fotoAbsensiUrl
                  }
                  title={
                    isLoadingFoto
                      ? "Mengecek foto..."
                      : fotoAbsensiUrl
                      ? "Foto sudah diupload"
                      : "Upload foto kelas"
                  }
                  className={`px-4 py-2 rounded-lg font-semibold shadow-md transition-colors ${
                    isUploadingFoto || isLoadingFoto || !!fotoAbsensiUrl
                      ? "bg-gray-400 cursor-not-allowed text-white"
                      : "bg-purple-600 hover:bg-purple-700 text-white"
                  }`}
                >
                  {isUploadingFoto ? "📤 Mengirim..." : "📷 Foto Kelas"}
                </button>
              </div>
            )}

          {/* TAMBAHKAN MODAL PREVIEW FOTO */}
          {previewFoto && (
            <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
              <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-auto">
                <div className="p-6">
                  <h3 className="text-xl font-bold text-gray-800 mb-4 text-center">
                    📷 Preview Foto Absensi
                  </h3>

                  <div className="mb-4 flex justify-center">
                    <img
                      src={previewFoto}
                      alt="Preview Foto Absensi"
                      className="rounded-lg shadow-md max-w-full max-h-96 object-contain"
                    />
                  </div>

                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 mb-4">
                    <p className="text-sm text-blue-700 text-center">
                      <strong>Kelas:</strong> {selectedKelas} •{" "}
                      <strong>Tanggal:</strong> {formatDateDDMMYYYY(date)}
                    </p>
                  </div>

                  <div className="flex gap-3 justify-center">
                    <button
                      onClick={cancelUploadFoto}
                      disabled={isUploadingFoto}
                      className="px-6 py-3 bg-gray-500 hover:bg-gray-600 text-white font-semibold rounded-lg shadow-md transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
                    >
                      ❌ Batal
                    </button>
                    <button
                      onClick={confirmUploadFoto}
                      disabled={isUploadingFoto}
                      className="px-6 py-3 bg-green-500 hover:bg-green-600 text-white font-semibold rounded-lg shadow-md transition-colors disabled:bg-gray-400 disabled:cursor-not-allowed"
                    >
                      {isUploadingFoto ? "⏳ Mengirim..." : "✅ Kirim Foto"}
                    </button>
                  </div>
                </div>
              </div>
            </div>
          )}

          <div className="text-center">
            <button
              onClick={() => setShowDebugInfo(!showDebugInfo)}
              className="text-sm bg-gray-200 hover:bg-gray-300 px-1 py-0.5 rounded-lg"
            >
              🔍 Info Debug
            </button>
          </div>
        </div>

        {/* Tampilan foto yang sudah diupload */}
        {fotoAbsensiUrl && (
          <div className="mb-6 bg-purple-50 border border-purple-200 rounded-lg p-4">
            <h4 className="font-semibold text-purple-700 mb-2 text-center">
              📷 Foto Absensi Kelas {selectedKelas} - {formatDateDDMMYYYY(date)}
            </h4>
            <div className="flex justify-center mb-3">
              <img
                src={fotoAbsensiUrl}
                alt="Foto Absensi Kelas"
                className="rounded-lg shadow-md max-w-full max-h-96 object-contain cursor-pointer hover:opacity-90 transition-opacity"
                onClick={() => setShowFotoFullscreen(true)}
                title="Klik untuk melihat ukuran penuh"
                onError={(e) => {
                  const target = e.target as HTMLImageElement;
                  target.style.display = "none";
                }}
              />
            </div>

            {/* Tombol Lihat & Hapus */}
            <div className="flex gap-3 justify-center">
              <button
                onClick={() => setShowFotoFullscreen(true)}
                className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-semibold shadow-md transition-colors"
              >
                🔍 Lihat Foto Besar
              </button>
              <button
                onClick={handleDeleteFoto}
                disabled={isDeletingFoto}
                className={`px-4 py-2 rounded-lg font-semibold shadow-md transition-colors ${
                  isDeletingFoto
                    ? "bg-gray-400 cursor-not-allowed text-white"
                    : "bg-red-600 hover:bg-red-700 text-white"
                }`}
              >
                {isDeletingFoto ? "⏳ Menghapus..." : "🗑️ Hapus Foto"}
              </button>
            </div>
          </div>
        )}

        {/* QR Scanner Section */}
        <div className="mb-6">
          <div className="text-center mb-4">
            {!isSunday(date) &&
              !isTanggalMerah(date) &&
              !isLiburSemester(date) && (
                <>
                  {!scanMode ? (
                    <button
                      onClick={startScanner}
                      disabled={isSendingData}
                      className={`px-6 py-3 rounded-lg font-bold shadow-md transition-colors text-lg ${
                        isSendingData
                          ? "bg-gray-400 cursor-not-allowed"
                          : "bg-blue-600 hover:bg-blue-700"
                      } text-white`}
                    >
                      📷 Mulai Scan QR Code
                    </button>
                  ) : (
                    <button
                      onClick={stopScanner}
                      disabled={isSendingData}
                      className={`px-6 py-3 rounded-lg font-bold shadow-md transition-colors text-lg ${
                        isSendingData
                          ? "bg-gray-400 cursor-not-allowed"
                          : "bg-red-600 hover:bg-red-700"
                      } text-white`}
                    >
                      ⏹️ Hentikan Scanner
                    </button>
                  )}
                </>
              )}
          </div>

          {isSendingData && (
            <div className="mb-4 bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
              <div className="flex items-center justify-center gap-2">
                <svg
                  className="animate-spin h-5 w-5 text-blue-600"
                  xmlns="http://www.w3.org/2000/svg"
                  fill="none"
                  viewBox="0 0 24 24"
                >
                  <circle
                    className="opacity-25"
                    cx="12"
                    cy="12"
                    r="10"
                    stroke="currentColor"
                    strokeWidth="4"
                  ></circle>
                  <path
                    className="opacity-75"
                    fill="currentColor"
                    d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                  ></path>
                </svg>
                <span className="text-blue-700 font-semibold">
                  💾 Menyimpan data ke server...
                </span>
              </div>
            </div>
          )}

          {scanMode && (
            <div className="relative border-4 border-blue-500 rounded-lg overflow-hidden">
              <video
                ref={videoRef}
                className="w-full h-96 object-cover bg-black"
                autoPlay
                playsInline
              />
              <div className="absolute top-0 left-0 right-0 bg-blue-600 text-white text-center py-2 font-semibold">
                📱 Arahkan kamera ke QR Code siswa
              </div>
              <div className="absolute bottom-0 left-0 right-0 bg-gradient-to-t from-black to-transparent p-4">
                <p className="text-white text-center text-sm">
                  Sudah Scan: {scannedStudents.size} / {filteredStudents.length}{" "}
                  siswa
                </p>
              </div>
            </div>
          )}

          {scannerError && (
            <div className="mt-4 bg-red-50 border border-red-200 rounded-lg p-4 text-center">
              <p className="text-red-600 font-semibold">❌ {scannerError}</p>
            </div>
          )}

          {scanMode && (
            <div className="mt-4 bg-blue-50 border border-blue-200 rounded-lg p-4">
              <h4 className="font-semibold text-blue-700 mb-2">
                📌 Cara Scan:
              </h4>
              <ol className="list-decimal list-inside text-sm text-blue-600 space-y-1">
                <li>Pastikan QR Code siswa terlihat jelas di kamera</li>
                <li>Tunggu hingga suara "beep" dan notifikasi muncul</li>
                <li>Data akan otomatis tersimpan ke server</li>
                <li>Siswa yang sudah scan tidak bisa scan ulang</li>
              </ol>
            </div>
          )}
        </div>

        {/* Galeri foto siswa yang sudah scan */}
        {scannedStudentPhotos.length > 0 && (
          <div className="mt-4">
            <h4 className="font-semibold text-gray-700 mb-3 text-center">
              ✅ Sudah Hadir ({scannedStudentPhotos.length} siswa)
            </h4>
            <div className="grid grid-cols-3 sm:grid-cols-4 md:grid-cols-5 gap-4">
              {scannedStudentPhotos.map((item, index) => (
                <div key={index} className="flex flex-col items-center gap-1">
                  {item.student.foto ? (
                    <img
                      src={getFotoUrlAttendance(item.student.foto)}
                      alt={item.student.name}
                      className="object-cover shadow-md rounded-lg"
                      style={{
                        width: "113px",
                        height: "85px",
                        border: "3px solid #4ade80",
                        borderRadius: "8px",
                      }}
                      onError={(e) => {
                        const target = e.target as HTMLImageElement;
                        target.style.display = "none";
                        if (target.parentElement) {
                          const div = document.createElement("div");
                          div.style.cssText =
                            "width:113px;height:85px;background:#dcfce7;border:3px solid #4ade80;border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:2rem;";
                          div.innerHTML = "👤";
                          target.parentElement.insertBefore(div, target);
                        }
                      }}
                    />
                  ) : (
                    <div
                      className="flex items-center justify-center shadow-md text-3xl"
                      style={{
                        width: "113px",
                        height: "85px",
                        background: "#dcfce7",
                        border: "3px solid #4ade80",
                        borderRadius: "8px",
                      }}
                    >
                      👤
                    </div>
                  )}
                  <p
                    className="text-xs text-center text-gray-700 font-medium leading-tight truncate"
                    style={{ maxWidth: "113px" }}
                  >
                    {item.student.name.split(" ")[0]}
                  </p>
                  <p className="text-xs text-green-600 font-semibold">
                    {item.timestamp.toLocaleTimeString("id-ID", {
                      hour: "2-digit",
                      minute: "2-digit",
                    })}
                  </p>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Modal Fullscreen Foto */}
        {showFotoFullscreen && fotoAbsensiUrl && (
          <div
            className="fixed inset-0 bg-black bg-opacity-90 flex items-center justify-center z-50 p-4"
            onClick={() => setShowFotoFullscreen(false)}
          >
            <div
              className="relative max-w-7xl max-h-full"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Tombol Close */}
              <button
                onClick={() => setShowFotoFullscreen(false)}
                className="absolute top-4 right-4 bg-red-600 hover:bg-red-700 text-white rounded-full w-12 h-12 flex items-center justify-center text-2xl font-bold shadow-lg z-10"
                title="Tutup"
              >
                ×
              </button>

              {/* Info */}
              <div className="absolute top-4 left-4 bg-black bg-opacity-70 text-white px-4 py-2 rounded-lg z-10">
                <p className="font-semibold">
                  Kelas {selectedKelas} • {formatDateDDMMYYYY(date)}
                </p>
              </div>

              {/* Foto Besar */}
              <img
                src={fotoAbsensiUrl}
                alt="Foto Absensi Fullscreen"
                className="rounded-lg shadow-2xl max-w-full max-h-[90vh] object-contain"
                onError={(e) => {
                  const target = e.target as HTMLImageElement;
                  target.style.display = "none";
                  alert("❌ Gagal memuat foto");
                  setShowFotoFullscreen(false);
                }}
              />

              {/* Tombol Download */}
              <div className="absolute bottom-4 left-1/2 transform -translate-x-1/2">
                <a
                  href={fotoAbsensiUrl}
                  download={`Foto_Absensi_${selectedKelas}_${formatDateDDMMYYYY(
                    date
                  )}.jpg`}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="bg-green-600 hover:bg-green-700 text-white px-6 py-3 rounded-lg font-semibold shadow-lg inline-block"
                >
                  📥 Download Foto
                </a>
              </div>
            </div>
          </div>
        )}

        {showDebugInfo && (
          <div className="mb-6 bg-yellow-50 border border-yellow-200 rounded-lg p-4">
            <h4 className="font-semibold text-yellow-800 mb-2">
              Informasi Debug:
            </h4>
            <div className="text-sm text-yellow-700 space-y-1">
              <p>
                <strong>Total Siswa:</strong> {students.length}
              </p>
              <p>
                <strong>Kelas yang Tersedia:</strong> {uniqueClasses.join(", ")}
              </p>
              <p>
                <strong>Kelas Terpilih:</strong> {selectedKelas}
              </p>
              <p>
                <strong>Siswa Terfilter:</strong> {filteredStudents.length}
              </p>
              <p>
                <strong>Siswa dengan Data Existing:</strong>{" "}
                {existingStudentIds.size}
              </p>
              <p>
                <strong>Siswa Sudah Scan QR:</strong> {scannedStudents.size}
              </p>
              <p>
                <strong>Semua Siswa Sudah Diabsen:</strong>{" "}
                {allStudentsHaveData ? "Ya" : "Tidak"}
              </p>
            </div>
          </div>
        )}

        <div className="mb-4 text-center">
          <p className="text-sm text-gray-600">
            Menampilkan: <span className="font-semibold">{selectedKelas}</span>{" "}
            • Tanggal:{" "}
            <span className="font-semibold">{formatDateDDMMYYYY(date)}</span> •
            Total Siswa:{" "}
            <span className="font-semibold">{filteredStudents.length}</span>
          </p>
        </div>

        {isSunday(date) ? (
          <div className="text-center py-12">
            <div className="bg-red-50 border-2 border-red-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🌅</div>
              <h3 className="text-2xl font-bold text-red-700 mb-2">
                Hari Minggu
              </h3>
              <p className="text-red-600">
                Tanggal {formatDateDDMMYYYY(date)} adalah hari Minggu.
              </p>
              <p className="text-sm text-red-500 mt-2">
                Tidak ada kegiatan belajar mengajar.
              </p>
            </div>
          </div>
        ) : isLiburSemester(date) ? (
          <div className="text-center py-12">
            <div className="bg-green-50 border-2 border-green-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🏖️</div>
              <h3 className="text-2xl font-bold text-green-700 mb-2">
                Libur Semester
              </h3>
              <p className="text-green-600">
                {getTanggalMerahInfo(date)?.deskripsi}
              </p>
              <p className="text-sm text-green-500 mt-2">
                Tanggal: {formatDateDDMMYYYY(date)}
              </p>
            </div>
          </div>
        ) : isTanggalMerah(date) ? (
          <div className="text-center py-12">
            <div className="bg-yellow-50 border-2 border-yellow-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">🎉</div>
              <h3 className="text-2xl font-bold text-yellow-700 mb-2">
                Hari Libur Nasional
              </h3>
              <p className="text-yellow-600 font-semibold text-lg">
                {getTanggalMerahInfo(date)?.deskripsi}
              </p>
              <p className="text-sm text-yellow-500 mt-2">
                Tanggal: {formatDateDDMMYYYY(date)}
              </p>
            </div>
          </div>
        ) : isBukanJadwalMengajar(date) ? (
          <div className="text-center py-12">
            <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-8 max-w-md mx-auto">
              <div className="text-6xl mb-4">📚</div>
              <h3 className="text-2xl font-bold text-blue-700 mb-2">
                Bukan Jadwal Mengajar
              </h3>
              <p className="text-blue-600">
                Hari ini ({formatDateDDMMYYYY(date)}) bukan jadwal mengajar Anda
                untuk kelas {selectedKelas}.
              </p>
              <p className="text-sm text-blue-500 mt-2">
                Silakan pilih tanggal atau kelas lain.
              </p>
            </div>
          </div>
        ) : (
          <>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
              <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
                <div className="text-green-600 font-bold text-lg">
                  {attendanceSummary.Hadir}
                </div>
                <div className="text-green-700 text-sm">Hadir</div>
              </div>
              <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
                <div className="text-yellow-600 font-bold text-lg">
                  {attendanceSummary.Izin}
                </div>
                <div className="text-yellow-700 text-sm">Izin</div>
              </div>
              <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
                <div className="text-blue-600 font-bold text-lg">
                  {attendanceSummary.Sakit}
                </div>
                <div className="text-blue-700 text-sm">Sakit</div>
              </div>
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
                <div className="text-red-600 font-bold text-lg">
                  {attendanceSummary.Alpha}
                </div>
                <div className="text-red-700 text-sm">Alpha</div>
              </div>
            </div>

            {isLoadingExistingData && (
              <div className="mb-4 bg-blue-50 border border-blue-200 rounded-lg p-4 text-center">
                <div className="flex items-center justify-center gap-2">
                  <svg
                    className="animate-spin h-5 w-5 text-blue-600"
                    xmlns="http://www.w3.org/2000/svg"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    ></circle>
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    ></path>
                  </svg>
                  <span className="text-blue-700 font-semibold">
                    ⏳ Mohon tunggu, sedang memuat data absensi...
                  </span>
                </div>
              </div>
            )}

            <div className="space-y-4 mb-6 overflow-x-auto">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-1 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-2 py-1 text-left text-sm font-semibold text-gray-700">
                      Nama Siswa
                    </th>

                    <th className="border border-gray-200 px-2 py-1 text-center text-sm font-semibold text-gray-700">
                      Status
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredStudents.map((s, index) => {
                    const isExisting = existingStudentIds.has(s.id);
                    const isScanned = scannedStudents.has(s.id);
                    return (
                      <tr
                        key={s.id}
                        className={`border-b border-gray-200 ${
                          isScanned ? "bg-green-50" : ""
                        }`}
                      >
                        <td
                          style={{ width: "1cm" }}
                          className="p-2 text-center"
                        >
                          <span className="text-sm font-medium text-gray-800">
                            {index + 1}
                          </span>
                        </td>
                        <td style={{ width: "5.5cm" }} className="p-2">
                          <p className="text-base font-semibold text-gray-800">
                            {s.name || "N/A"}
                            {isScanned && (
                              <span className="ml-2 text-xs bg-green-500 text-white px-2 py-1 rounded">
                                ✓ Tersimpan
                              </span>
                            )}
                            {isExisting && !isScanned && (
                              <span className="ml-2 text-xs bg-blue-500 text-white px-2 py-1 rounded">
                                ✓ Sudah Ada
                              </span>
                            )}
                          </p>
                          <p className="text-sm text-gray-500">
                            Kelas {s.kelas || "N/A"} • NISN: {s.nisn || "N/A"}
                          </p>
                        </td>

                        <td style={{ width: "5cm" }} className="p-2">
                          <div className="flex flex-col gap-2">
                            {/* Tombol Status */}
                            <div className="flex justify-between">
                              {(
                                ["Hadir", "Izin", "Sakit", "Alpha"] as const
                              ).map((status) => (
                                <button
                                  key={status}
                                  onClick={() => setStatus(s.id, status)}
                                  style={{ width: "1cm" }}
                                  className={`px-1 py-0.5 rounded-lg text-xs font-medium transition-colors ${
                                    attendance[date]?.[s.id] === status &&
                                    attendance[date]?.[s.id] !== ""
                                      ? `${statusColor[status]} text-white opacity-70`
                                      : "bg-gray-200 text-gray-500 cursor-not-allowed"
                                  }`}
                                  disabled={true}
                                >
                                  {status}
                                </button>
                              ))}
                            </div>

                            {/* TAMBAHKAN: Input Keterangan */}
                            <input
                              type="text"
                              placeholder="Keterangan (opsional)"
                              value={
                                (keterangan[date] && keterangan[date][s.id]) ||
                                ""
                              }
                              onChange={(e) =>
                                setKeteranganValue(s.id, e.target.value)
                              }
                              disabled={true}
                              className="w-full px-2 py-1 text-xs border rounded bg-gray-100 text-gray-500 cursor-not-allowed"
                            />

                            {isExisting && (
                              <p className="text-xs text-gray-400 mt-1">
                                🕐{" "}
                                {existingAttendanceData.find(
                                  (r: any) => r.nama === s.name
                                )?.jam || "-"}
                              </p>
                            )}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const MonthlyRecapTab: React.FC<{
  onRefresh: () => void;
  uniqueClasses: string[];
  students: Student[];
}> = ({ onRefresh, uniqueClasses, students }) => {
  const [recapData, setRecapData] = useState<MonthlyRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>("semua");

  const months = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
  ] as const;

  // Dapatkan bulan berjalan secara otomatis
  const getCurrentMonth = () => {
    const currentMonthIndex = new Date().getMonth(); // 0-11
    return months[currentMonthIndex];
  };

  const [selectedBulan, setSelectedBulan] = useState<string>(getCurrentMonth());
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );
  const [loading, setLoading] = useState<boolean>(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);

  useEffect(() => {
    setLoading(true);
    console.log(
      "Mengambil data rekap dengan kelas:",
      selectedKelas,
      "dan bulan:",
      selectedBulan
    );
    fetch(
      `${endpoint}?action=monthlyRecap&kelas=${
        selectedKelas === "Semua" ? "" : selectedKelas
      }&bulan=${selectedBulan.toLowerCase()}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        console.log("Respons data rekap:", data);
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert("❌ Gagal memuat data rekap: " + data.message);
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data rekap. Cek console untuk detail.");
        setRecapData([]);
        setLoading(false);
      });

    // Fetch school data
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedBulan, onRefresh]);

  const filteredRecapData = React.useMemo(() => {
    if (selectedKelas === "Semua") {
      return recapData;
    }
    console.log("Menyaring data rekap untuk kelas:", selectedKelas);
    return recapData.filter((item) => {
      const itemKelas = String(item.kelas).trim();
      const result = itemKelas === selectedKelas;
      console.log("Kelas item:", itemKelas, "cocok?", result);
      return result;
    });
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): StatusSummary => {
    const summary: StatusSummary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();

  const downloadExcel = () => {
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item, index) => [
        index + 1, // Nomor urut
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "",
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "",
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [
      { wch: 5 }, // Lebar kolom No. (sempit)
      { wch: 25 }, // Nama
      { wch: 10 }, // Kelas
      { wch: 10 }, // Hadir
      { wch: 10 }, // Alpha
      { wch: 10 }, // Izin
      { wch: 10 }, // Sakit
      { wch: 10 }, // % Hadir
    ];
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Bulanan");

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    // Title - Format sama dengan Daftar Hadir
    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";

    // ✅ TAMBAHKAN: Ambil tahun dari selectedDate
    const tahunDariTanggal = new Date(selectedDate).getFullYear();

    // Judul dalam 1 baris - gunakan tahunDariTanggal
    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas}  ${namaSekolah}  ${selectedBulan.toUpperCase()} ${tahunDariTanggal}`;

    doc.setFontSize(12); // Ukuran font sama dengan daftar hadir
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });

    currentY += 10;

    // Table headers and data
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item, index) => [
      index + 1, // Nomor urut
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "",
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];

    const percentRow = [
      "",
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 10 }, // No. (sempit)
        1: { cellWidth: 50 }, // Nama
        2: { cellWidth: 20 }, // Kelas
        3: { cellWidth: 20 }, // Hadir
        4: { cellWidth: 20 }, // Alpha
        5: { cellWidth: 20 }, // Izin
        6: { cellWidth: 20 }, // Sakit
        7: { cellWidth: 20 }, // % Hadir
      },
    });

    // Update currentY after the table
    currentY = (doc as any).lastAutoTable.finalY + 10;

    // ✅ TAMBAHAN BARU: Tabel Informasi Jumlah Siswa
    const pageHeight = doc.internal.pageSize.getHeight();
    const bottomMargin = 20;
    const spaceNeededForStudentTable = 20;
    const spaceNeededForSignatures = 60; // Ruang untuk tanda tangan

    // ✅ CEK APAKAH TANDA TANGAN + TABEL SISWA MUAT DI HALAMAN INI
    // Cek apakah ada cukup ruang untuk tabel jumlah siswa DAN tanda tangan
    if (
      currentY + spaceNeededForStudentTable + spaceNeededForSignatures >
      pageHeight - bottomMargin
    ) {
      doc.addPage();
      currentY = margin;
    }

    // Hitung jumlah siswa berdasarkan jenis kelamin
    const genderSummary = filteredRecapData.reduce(
      (acc, student) => {
        // Ambil data siswa dari filteredRecapData untuk mendapatkan jenis kelamin
        const studentData = students.find((s) => s.name === student.nama);
        if (studentData) {
          const jenisKelamin = String(studentData.jenisKelamin || "")
            .trim()
            .toUpperCase();
          if (jenisKelamin === "L" || jenisKelamin === "LAKI-LAKI") {
            acc.lakiLaki++;
          } else if (jenisKelamin === "P" || jenisKelamin === "PEREMPUAN") {
            acc.perempuan++;
          }
        }
        return acc;
      },
      { lakiLaki: 0, perempuan: 0 }
    );

    const totalSiswa = genderSummary.lakiLaki + genderSummary.perempuan;

    doc.setFontSize(10);
    doc.setFont("Times", "bold");
    doc.text("JUMLAH SISWA", margin, currentY, { align: "left" });
    currentY += 3;

    const tableWidth = (pageWidth - 2 * margin) * 0.4;

    autoTable(doc, {
      head: [["LAKI-LAKI", "PEREMPUAN", "TOTAL SISWA"]],
      body: [
        [
          genderSummary.lakiLaki.toString(),
          genderSummary.perempuan.toString(),
          totalSiswa.toString(),
        ],
      ],
      startY: currentY,
      margin: { left: margin, right: pageWidth - margin - tableWidth },
      tableWidth: tableWidth,
      theme: "grid",
      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1,
        halign: "center",
        valign: "middle",
        lineWidth: 0.5,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineWidth: 1,
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 10,
        lineWidth: 1,
      },
      columnStyles: {
        0: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        1: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        2: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    // Add school data (Principal and Teacher details)
    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      // Add place and date above Guru Kelas, centered
      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${
        schoolData.namaKota || "Makassar"
      }, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50; // Signature width is 50
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5; // Keep close to "Guru Kelas"

      // Principal Section
      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      // Calculate width for signatures
      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas (lebar lebih besar untuk tanda tangan panjang)
          canvas.height = 50; // Sesuaikan ukuran canvas (tinggi cukup untuk garis tanda tangan)
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdKepsek); // schoolData.ttdKepsek adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            leftColumnX + 10,
            currentY - 3,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Kepsek signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Kepala Sekolah.",
            leftColumnX + 10,
            currentY - 3 + 10
          );
        }
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas
          canvas.height = 50;
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdGuru); // schoolData.ttdGuru adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            rightColumnX + 10,
            currentY - 5,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Guru signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Guru.",
            rightColumnX + 10,
            currentY - 5 + 10
          );
        }
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text(
        `${schoolData.statusGuru || "Guru Kelas"},`,
        rightColumnX + 25,
        currentY - 2,
        {
          align: "center",
        }
      );

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Bulanan_${selectedBulan}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Bulanan
        </h2>
        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => {
                console.log("Mengubah filter kelas ke:", e.target.value);
                setSelectedKelas(e.target.value);
              }}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Bulan</p>
            <select
              value={selectedBulan}
              onChange={(e) => setSelectedBulan(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {months.map((month) => (
                <option key={month} value={month}>
                  {month}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Bulanan pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk {selectedBulan} kelas {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau bulan lain.
            </p>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-0.5 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-2 py-0.5 text-center text-sm text-gray-600 font-medium">
                        {index + 1}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

const GraphTab: React.FC<{
  uniqueClasses: string[];
}> = ({ uniqueClasses }) => {
  const [graphData, setGraphData] = useState<GraphData>({
    Januari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Februari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Maret: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    April: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Mei: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juni: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Juli: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Agustus: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    September: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Oktober: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    November: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
    Desember: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
  });
  const [selectedKelas, setSelectedKelas] = useState<string>(
    uniqueClasses.length > 0 ? uniqueClasses[0] : "Tidak Ada"
  );
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("2");
  const [statusVisibility, setStatusVisibility] = useState<StatusVisibility>({
    Hadir: true,
    Alpha: true,
    Izin: true,
    Sakit: true,
  });
  const [loading, setLoading] = useState<boolean>(true);

  const uniqueClassesWithDefault = React.useMemo(() => {
    return ["Tidak Ada", ...uniqueClasses.filter((kelas) => kelas !== "Semua")];
  }, [uniqueClasses]);

  useEffect(() => {
    setLoading(true);
    fetch(
      `${endpoint}?action=graphData&kelas=${
        selectedKelas === "Tidak Ada" ? "" : selectedKelas
      }&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setGraphData(data.data || {});
        } else {
          alert("❌ Gagal memuat data grafik: " + data.message);
          setGraphData({
            Januari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Februari: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Maret: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            April: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Mei: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Juni: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Juli: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Agustus: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            September: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Oktober: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            November: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
            Desember: { Hadir: 0, Alpha: 0, Izin: 0, Sakit: 0 },
          });
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal memuat data grafik. Cek console untuk detail.");
        setLoading(false);
      });
  }, [selectedKelas, selectedSemester]);

  const semesterMonths: Record<"1" | "2", string[]> = {
    "1": ["Juli", "Agustus", "September", "Oktober", "November", "Desember"],
    "2": ["Januari", "Februari", "Maret", "April", "Mei", "Juni"],
  };

  const chartData: ChartData<"bar", number[], string> = {
    labels: semesterMonths[selectedSemester],
    datasets: [
      ...(statusVisibility.Hadir
        ? [
            {
              label: "Hadir",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Hadir || 0
              ),
              backgroundColor: "rgba(75, 192, 192, 0.6)",
              borderColor: "rgba(75, 192, 192, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Alpha
        ? [
            {
              label: "Alpha",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Alpha || 0
              ),
              backgroundColor: "rgba(255, 99, 132, 0.6)",
              borderColor: "rgba(255, 99, 132, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Izin
        ? [
            {
              label: "Izin",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Izin || 0
              ),
              backgroundColor: "rgba(255, 205, 86, 0.6)",
              borderColor: "rgba(255, 205, 86, 1)",
              borderWidth: 1,
            },
          ]
        : []),
      ...(statusVisibility.Sakit
        ? [
            {
              label: "Sakit",
              data: semesterMonths[selectedSemester].map(
                (month: string) => graphData[month]?.Sakit || 0
              ),
              backgroundColor: "rgba(54, 162, 235, 0.6)",
              borderColor: "rgba(54, 162, 235, 1)",
              borderWidth: 1,
            },
          ]
        : []),
    ],
  };

  const chartOptions: ChartOptions<"bar"> = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        position: "top" as const,
        onClick: (
          e: ChartEvent,
          legendItem: LegendItem,
          legend: {
            chart: {
              data: { datasets: { hidden?: boolean }[] };
              update: () => void;
            };
          }
        ) => {
          const index = legendItem.datasetIndex;
          if (index !== undefined) {
            const ci = legend.chart.data.datasets[index];
            ci.hidden = !ci.hidden;
            legend.chart.update();
            setStatusVisibility((prev) => ({
              ...prev,
              [legendItem.text as keyof StatusVisibility]: !ci.hidden,
            }));
          }
        },
      },
      title: {
        display: true,
        text: `Persentase Kehadiran Kelas ${selectedKelas} Semester ${selectedSemester} 2025`,
      },
    },
    scales: {
      y: {
        beginAtZero: true,
        max: 100,
        ticks: {
          stepSize: 10,
          font: {
            size: 10,
          },
          autoSkip: false,
          maxTicksLimit: 11,
        },
        title: { display: true, text: "Persentase (%)" },
      },
      x: {
        ticks: {
          font: {
            size: 10,
          },
        },
      },
    },
  };

  const handleStatusToggle = (status: keyof StatusVisibility) => {
    setStatusVisibility((prev) => ({
      ...prev,
      [status]: !prev[status],
    }));
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📈 Grafik Kehadiran
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClassesWithDefault.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1 (Juli-Des)</option>
              <option value="2">Semester 2 (Jan-Jun)</option>
            </select>
          </div>
        </div>

        <div className="mb-6 flex flex-wrap gap-4 justify-center">
          {(["Hadir", "Alpha", "Izin", "Sakit"] as const).map((status) => (
            <label key={status} className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={statusVisibility[status]}
                onChange={() => handleStatusToggle(status)}
                className="h-4 w-4 text-blue-600 rounded"
              />
              <span className="text-sm text-gray-700">{status}</span>
            </label>
          ))}
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat grafik...</p>
          </div>
        ) : selectedKelas === "Tidak Ada" ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Tidak ada data untuk ditampilkan.</p>
          </div>
        ) : (
          <div
            className="h-96"
            style={{
              minHeight: "300px",
              maxHeight: "500px",
            }}
          >
            <Bar data={chartData} options={chartOptions} />
          </div>
        )}
      </div>
    </div>
  );
};

const SemesterRecapTab: React.FC<{
  uniqueClasses: string[];
  students: Student[]; // ✅ TAMBAHKAN INI
}> = ({
  uniqueClasses,
  students, // ✅ DAN INI
}) => {
  const [recapData, setRecapData] = useState<SemesterRecap[]>([]);
  const [selectedKelas, setSelectedKelas] = useState<string>("Semua");
  const [selectedSemester, setSelectedSemester] = useState<"1" | "2">("1");
  const [selectedDate, setSelectedDate] = useState<string>(
    new Date().toISOString().split("T")[0]
  );

  const [loading, setLoading] = useState<boolean>(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);

  useEffect(() => {
    setLoading(true);
    const sheetName =
      selectedSemester === "1" ? SHEET_SEMESTER1 : SHEET_SEMESTER2;
    fetch(
      `${endpoint}?action=semesterRecap&kelas=${
        selectedKelas === "Semua" ? "" : selectedKelas
      }&semester=${selectedSemester}`
    )
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success) {
          setRecapData(data.data || []);
        } else {
          alert(
            `❌ Gagal memuat data rekap ${
              selectedSemester === "1" ? "Semester 1" : "Semester 2"
            }: ${data.message}`
          );
          setRecapData([]);
        }
        setLoading(false);
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert(
          `❌ Gagal memuat data rekap ${
            selectedSemester === "1" ? "Semester 1" : "Semester 2"
          }. Cek console untuk detail.`
        );
        setRecapData([]);
        setLoading(false);
      });

    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
        alert("❌ Gagal memuat data sekolah. Cek console untuk detail.");
      });
  }, [selectedKelas, selectedSemester]);

  const filteredRecapData = React.useMemo(() => {
    if (selectedKelas === "Semua") {
      return recapData;
    }
    return recapData.filter(
      (item) => String(item.kelas).trim() === selectedKelas
    );
  }, [recapData, selectedKelas]);

  const getStatusSummary = (): {
    Hadir: number;
    Izin: number;
    Sakit: number;
    Alpha: number;
  } => {
    const summary = { Hadir: 0, Izin: 0, Sakit: 0, Alpha: 0 };
    filteredRecapData.forEach((item) => {
      summary.Hadir += item.hadir || 0;
      summary.Alpha += item.alpa || 0;
      summary.Izin += item.izin || 0;
      summary.Sakit += item.sakit || 0;
    });
    return summary;
  };

  const statusSummary = getStatusSummary();

  const downloadExcel = () => {
    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const data = [
      headers,
      ...filteredRecapData.map((item, index) => [
        index + 1, // Nomor urut
        item.nama || "N/A",
        item.kelas || "N/A",
        item.hadir || 0,
        item.alpa || 0,
        item.izin || 0,
        item.sakit || 0,
        item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
      ]),
      [
        "",
        "TOTAL",
        "",
        statusSummary.Hadir,
        statusSummary.Alpha,
        statusSummary.Izin,
        statusSummary.Sakit,
        "",
      ],
      [
        "",
        "PERSEN",
        "",
        `${(
          (statusSummary.Hadir /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Alpha /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Izin /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        `${(
          (statusSummary.Sakit /
            (statusSummary.Hadir +
              statusSummary.Alpha +
              statusSummary.Izin +
              statusSummary.Sakit)) *
          100
        ).toFixed(2)}%`,
        "",
      ],
    ];

    const ws = XLSX.utils.aoa_to_sheet(data);
    ws["!cols"] = [
      { wch: 5 }, // Lebar kolom No. (sempit)
      { wch: 25 }, // Nama
      { wch: 10 }, // Kelas
      { wch: 10 }, // Hadir
      { wch: 10 }, // Alpha
      { wch: 10 }, // Izin
      { wch: 10 }, // Sakit
      { wch: 10 }, // % Hadir
    ];
    const headerStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "FFFF00" } },
      alignment: { horizontal: "center" },
    };
    const totalStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    const percentStyle = {
      font: { bold: true },
      fill: { fgColor: { rgb: "D3D3D3" } },
      alignment: { horizontal: "center" },
    };
    headers.forEach((header, index) => {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: index });
      ws[cellAddress] = { ...ws[cellAddress], s: headerStyle };
    });
    const totalRow = filteredRecapData.length + 1;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${totalRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: totalStyle };
    });
    const percentRow = filteredRecapData.length + 2;
    ["A", "B", "C", "D", "E", "F", "G", "H"].forEach((col, idx) => {
      const cellAddress = `${col}${percentRow}`;
      ws[cellAddress] = { ...ws[cellAddress], s: percentStyle };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Rekap Semester");
    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };

  const downloadPDF = async () => {
    const doc = new jsPDF();
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 14;
    const lineSpacing = 5;
    let currentY = margin;

    doc.setFont("Times", "roman");

    // Title - Format sama dengan Daftar Hadir
    const namaSekolah = schoolData?.namaSekolah || "UPT SDN 13 BATANG";

    // ✅ TAMBAHKAN: Ambil tahun dari selectedDate
    const tahunDariTanggal = new Date(selectedDate).getFullYear();

    // Judul dalam 1 baris dengan format Semester - gunakan tahunDariTanggal
    const semesterLabel =
      selectedSemester === "1" ? "SEMESTER 1" : "SEMESTER 2";
    const title = `REKAP ABSENSI SISWA KELAS ${selectedKelas}  ${namaSekolah}  ${semesterLabel} ${tahunDariTanggal}`;

    doc.setFontSize(12); // Ukuran font sama dengan daftar hadir
    doc.setFont("Times", "bold");
    doc.text(title, pageWidth / 2, currentY, { align: "center" });

    currentY += 10;

    const headers = [
      "No.",
      "Nama",
      "Kelas",
      "Hadir",
      "Alpha",
      "Izin",
      "Sakit",
      "% Hadir",
    ];
    const body = filteredRecapData.map((item, index) => [
      index + 1, // Nomor urut
      item.nama || "N/A",
      item.kelas || "N/A",
      item.hadir || 0,
      item.alpa || 0,
      item.izin || 0,
      item.sakit || 0,
      item.persenHadir !== undefined ? `${item.persenHadir}%` : "N/A",
    ]);

    const totalRow = [
      "",
      "TOTAL",
      "",
      statusSummary.Hadir,
      statusSummary.Alpha,
      statusSummary.Izin,
      statusSummary.Sakit,
      "",
    ];

    const percentRow = [
      "",
      "PERSEN",
      "",
      `${(
        (statusSummary.Hadir /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Alpha /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Izin /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      `${(
        (statusSummary.Sakit /
          (statusSummary.Hadir +
            statusSummary.Alpha +
            statusSummary.Izin +
            statusSummary.Sakit)) *
        100
      ).toFixed(2)}%`,
      "",
    ];

    autoTable(doc, {
      head: [headers],
      body: [...body, totalRow, percentRow],
      startY: currentY,
      styles: { font: "Times", fontSize: 8, cellPadding: 2 },
      headStyles: {
        fillColor: [255, 255, 0],
        textColor: [0, 0, 0],
        fontStyle: "bold",
      },
      alternateRowStyles: { fillColor: [240, 240, 240] },
      columnStyles: {
        0: { cellWidth: 10 }, // No. (sempit)
        1: { cellWidth: 50 }, // Nama
        2: { cellWidth: 20 }, // Kelas
        3: { cellWidth: 20 }, // Hadir
        4: { cellWidth: 20 }, // Alpha
        5: { cellWidth: 20 }, // Izin
        6: { cellWidth: 20 }, // Sakit
        7: { cellWidth: 20 }, // % Hadir
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    // ✅ TAMBAHAN BARU: Tabel Informasi Jumlah Siswa
    const pageHeight = doc.internal.pageSize.getHeight();
    const bottomMargin = 20;
    const spaceNeededForStudentTable = 20;
    const spaceNeededForSignatures = 60; // Ruang untuk tanda tangan

    // ✅ CEK APAKAH TANDA TANGAN + TABEL SISWA MUAT DI HALAMAN INI
    // Cek apakah ada cukup ruang untuk tabel jumlah siswa DAN tanda tangan
    if (
      currentY + spaceNeededForStudentTable + spaceNeededForSignatures >
      pageHeight - bottomMargin
    ) {
      doc.addPage();
      currentY = margin;
    }

    // Hitung jumlah siswa berdasarkan jenis kelamin dari filteredRecapData
    const genderSummary = filteredRecapData.reduce(
      (acc, student) => {
        // Cari data lengkap siswa untuk mendapatkan jenis kelamin
        const studentData = students.find((s) => s.name === student.nama);
        if (studentData) {
          const jenisKelamin = String(studentData.jenisKelamin || "")
            .trim()
            .toUpperCase();
          if (jenisKelamin === "L" || jenisKelamin === "LAKI-LAKI") {
            acc.lakiLaki++;
          } else if (jenisKelamin === "P" || jenisKelamin === "PEREMPUAN") {
            acc.perempuan++;
          }
        }
        return acc;
      },
      { lakiLaki: 0, perempuan: 0 }
    );

    const totalSiswa = genderSummary.lakiLaki + genderSummary.perempuan;

    doc.setFontSize(10);
    doc.setFont("Times", "bold");
    doc.text("JUMLAH SISWA", margin, currentY, { align: "left" });
    currentY += 3;

    const tableWidth = (pageWidth - 2 * margin) * 0.4;

    autoTable(doc, {
      head: [["LAKI-LAKI", "PEREMPUAN", "TOTAL SISWA"]],
      body: [
        [
          genderSummary.lakiLaki.toString(),
          genderSummary.perempuan.toString(),
          totalSiswa.toString(),
        ],
      ],
      startY: currentY,
      margin: { left: margin, right: pageWidth - margin - tableWidth },
      tableWidth: tableWidth,
      theme: "grid",
      styles: {
        font: "Times",
        fontSize: 7,
        cellPadding: 1,
        halign: "center",
        valign: "middle",
        lineWidth: 0.5,
      },
      headStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        lineWidth: 1,
      },
      bodyStyles: {
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        fontStyle: "bold",
        fontSize: 10,
        lineWidth: 1,
      },
      columnStyles: {
        0: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        1: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
        2: { cellWidth: tableWidth / 3, fillColor: [255, 255, 255] },
      },
    });

    currentY = (doc as any).lastAutoTable.finalY + 10;

    if (schoolData) {
      doc.setFontSize(10);
      doc.setFont("Times", "roman");

      const formattedDate = new Date(selectedDate).toLocaleDateString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
      });
      const placeDateText = `${
        schoolData.namaKota || "Makassar"
      }, ${formattedDate}`;
      const rightColumnX = pageWidth - margin - 50; // Signature width is 50
      doc.text(placeDateText, rightColumnX + 25, currentY - 1, {
        align: "center",
      });
      currentY += 5; // Keep close to "Guru Kelas"

      const principalText = [
        "Kepala Sekolah,",
        "",
        "",
        `( ${schoolData.namaKepsek || "N/A"} )`,
        `NIP: ${schoolData.nipKepsek || "N/A"}`,
      ];
      const teacherText = [
        "Guru Kelas,",
        "",
        "",
        `( ${schoolData.namaGuru || "N/A"} )`,
        `NIP: ${schoolData.nipGuru || "N/A"}`,
      ];

      const signatureWidth = 30;
      const signatureHeight = 20;
      const leftColumnX = margin;

      // Principal signature and text
      if (schoolData.ttdKepsek) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas (lebar lebih besar untuk tanda tangan panjang)
          canvas.height = 50; // Sesuaikan ukuran canvas (tinggi cukup untuk garis tanda tangan)
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdKepsek); // schoolData.ttdKepsek adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            leftColumnX + 10,
            currentY - 3,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Kepsek signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Kepala Sekolah.",
            leftColumnX + 10,
            currentY - 3 + 10
          );
        }
      }

      // Pisahkan "Kepala Sekolah" dengan posisi yang lebih tinggi
      doc.text("Kepala Sekolah,", leftColumnX + 25, currentY - 2, {
        align: "center",
      });

      // Kosong dan kosong
      doc.text("", leftColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", leftColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama kepala sekolah dengan format bold dan underline
      const principalName = schoolData.namaKepsek || "N/A";
      doc.setFont("Times", "bold");
      doc.text(principalName, leftColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to principal name
      const principalNameText = principalName;
      const textWidth = doc.getTextWidth(principalNameText);
      const textX = leftColumnX + 25 - textWidth / 2;
      doc.line(
        textX,
        currentY + 3.5 * lineSpacing + 1,
        textX + textWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipKepsek || "N/A"}`,
        leftColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );

      // Teacher signature and text
      if (schoolData.ttdGuru) {
        try {
          const canvas = document.createElement("canvas");
          canvas.width = 150; // Sesuaikan ukuran canvas
          canvas.height = 50;
          const ctx = canvas.getContext("2d");
          const v = await Canvg.from(ctx, schoolData.ttdGuru); // schoolData.ttdGuru adalah base64 SVG
          v.start();
          const pngData = canvas.toDataURL("image/png");
          doc.addImage(
            pngData,
            "PNG",
            rightColumnX + 10,
            currentY - 5,
            signatureWidth,
            signatureHeight
          ); // Sesuaikan posisi sesuai asli
        } catch (error) {
          console.error("Error rendering Guru signature:", error);
          doc.setFontSize(10);
          doc.text(
            "Gagal render tanda tangan Guru.",
            rightColumnX + 10,
            currentY - 5 + 10
          );
        }
      }

      // Pisahkan "Guru Kelas" dengan posisi yang lebih tinggi
      doc.text(
        `${schoolData.statusGuru || "Guru Kelas"},`,
        rightColumnX + 25,
        currentY - 2,
        {
          align: "center",
        }
      );

      // Kosong dan kosong
      doc.text("", rightColumnX + 25, currentY + lineSpacing, {
        align: "center",
      });
      doc.text("", rightColumnX + 25, currentY + 2 * lineSpacing, {
        align: "center",
      });

      // Nama guru dengan format bold dan underline
      const teacherName = schoolData.namaGuru || "N/A";
      doc.setFont("Times", "bold");
      doc.text(teacherName, rightColumnX + 25, currentY + 3.5 * lineSpacing, {
        align: "center",
      });

      // Add underline to teacher name
      const teacherNameText = teacherName;
      const teacherTextWidth = doc.getTextWidth(teacherNameText);
      const teacherTextX = rightColumnX + 25 - teacherTextWidth / 2;
      doc.line(
        teacherTextX,
        currentY + 3.5 * lineSpacing + 1,
        teacherTextX + teacherTextWidth,
        currentY + 3.5 * lineSpacing + 1
      );

      // Reset font and add NIP
      doc.setFont("Times", "roman");
      doc.text(
        `NIP. ${schoolData.nipGuru || "N/A"}`,
        rightColumnX + 25,
        currentY + 4.5 * lineSpacing,
        {
          align: "center",
        }
      );
    } else {
      doc.setFontSize(10);
      doc.text("Data sekolah tidak tersedia.", margin, currentY);
    }

    const date = new Date()
      .toLocaleString("id-ID", {
        day: "2-digit",
        month: "long",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      })
      .replace(/ /g, "_")
      .replace(/:/g, "-");
    const fileName = `Rekap_Semester_${selectedSemester}_${selectedKelas}_${date}.pdf`;
    doc.save(fileName);
  };

  return (
    <div className="max-w-4xl mx-auto" style={{ paddingBottom: "70px" }}>
      <div className="bg-white p-6 rounded-lg shadow-md">
        <h2 className="text-2xl font-bold text-center text-blue-700 mb-6">
          📊 Rekap Absensi Semester
        </h2>

        <div className="mb-6 flex flex-col md:flex-row gap-4 items-center justify-center">
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Filter Kelas</p>
            <select
              value={selectedKelas}
              onChange={(e) => setSelectedKelas(e.target.value)}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              {uniqueClasses.map((kelas) => (
                <option key={kelas} value={kelas}>
                  {kelas}
                </option>
              ))}
            </select>
          </div>
          <div className="text-center">
            <p className="text-sm text-gray-500 mb-2">Pilih Semester</p>
            <select
              value={selectedSemester}
              onChange={(e) => setSelectedSemester(e.target.value as "1" | "2")}
              className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
            >
              <option value="1">Semester 1</option>
              <option value="2">Semester 2</option>
            </select>
          </div>
        </div>

        {/* Separator line and PDF settings section */}
        <div className="border-t border-gray-200 pt-4 mb-6">
          <p className="text-center text-sm font-medium text-gray-700 mb-4">
            Pengaturan Tanggal & Nama Tempat <br /> untuk Rekap Semester pada
            File PDF
          </p>
          <div className="flex flex-col md:flex-row gap-4 items-center justify-center">
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-2">Pilih Tanggal</p>
              <input
                type="date"
                value={selectedDate}
                onChange={(e) => setSelectedDate(e.target.value)}
                className="border border-gray-300 rounded-lg px-1 py-0.5 shadow-sm bg-white min-w-32"
              />
            </div>
          </div>
        </div>

        <div className="grid grid-cols-2 md:grid-cols-4 gap-4 mb-6">
          <div className="bg-green-50 border border-green-200 rounded-lg p-3 text-center">
            <div className="text-green-600 font-bold text-lg">
              {statusSummary.Hadir}
            </div>
            <div className="text-green-700 text-sm">Hadir</div>
          </div>
          <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-3 text-center">
            <div className="text-yellow-600 font-bold text-lg">
              {statusSummary.Izin}
            </div>
            <div className="text-yellow-700 text-sm">Izin</div>
          </div>
          <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-center">
            <div className="text-blue-600 font-bold text-lg">
              {statusSummary.Sakit}
            </div>
            <div className="text-blue-700 text-sm">Sakit</div>
          </div>
          <div className="bg-red-50 border border-red-200 rounded-lg p-3 text-center">
            <div className="text-red-600 font-bold text-lg">
              {statusSummary.Alpha}
            </div>
            <div className="text-red-700 text-sm">Alpha</div>
          </div>
        </div>

        {loading ? (
          <div className="text-center py-8">
            <p className="text-gray-500">Memuat rekap...</p>
          </div>
        ) : filteredRecapData.length === 0 ? (
          <div className="text-center py-8">
            <p className="text-gray-500">
              Tidak ada data rekap untuk Semester {selectedSemester} kelas{" "}
              {selectedKelas}.
            </p>
            <p className="text-sm text-gray-400 mt-2">
              Coba pilih kelas atau semester lain.
            </p>
          </div>
        ) : (
          <>
            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse border border-gray-200">
                <thead>
                  <tr className="bg-gray-100">
                    <th className="border border-gray-200 px-2 py-0.5 text-center text-sm font-semibold text-gray-700">
                      No.
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Nama
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-left text-sm font-semibold text-gray-700">
                      Kelas
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Hadir
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Alpha
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Izin
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      Sakit
                    </th>
                    <th className="border border-gray-200 px-1 py-0.5 text-center text-sm font-semibold text-gray-700">
                      % Hadir
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecapData.map((item, index) => (
                    <tr
                      key={index}
                      className={index % 2 === 0 ? "bg-white" : "bg-gray-50"}
                    >
                      <td className="border border-gray-200 px-2 py-0.5 text-center text-sm text-gray-600 font-medium">
                        {index + 1}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.nama || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-sm text-gray-600">
                        {item.kelas || "N/A"}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.hadir || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.alpa || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.izin || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.sakit || 0}
                      </td>
                      <td className="border border-gray-200 px-1 py-0.5 text-center text-sm text-gray-600">
                        {item.persenHadir !== undefined
                          ? `${item.persenHadir}%`
                          : "N/A"}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="mt-6 flex gap-4 justify-center">
              <button
                onClick={downloadExcel}
                className="px-1 py-0.5 bg-green-600 hover:bg-green-700 text-white rounded-lg font-medium transition-colors"
              >
                📥 Download Excel
              </button>
              <button
                onClick={downloadPDF}
                className="px-1 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-medium transition-colors"
              >
                📄 Download PDF
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
};

// Komponen SplashScreen
const SplashScreen: React.FC = () => {
  return (
    <div className="fixed inset-0 bg-white flex flex-col items-center justify-center z-50">
      <style>
        {`
          @keyframes pulse {
            0% {
              transform: scale(1);
              opacity: 1;
            }
            50% {
              transform: scale(1.2);
              opacity: 0.7;
            }
            100% {
              transform: scale(1);
              opacity: 1;
            }
          }
          .animate-pulse-custom {
            animation: pulse 2s infinite;
          }
        `}
      </style>
      <img
        src="\images\logo_1.png"
        alt="Logo Aplikasi"
        className="w-52 h-70 mb-4 animate-pulse-custom" //Pengaturan ukuran logo
      />
      <p className="text-gray-800 text-lg font-semibold mt-6">Kelas 6A</p>
    </div>
  );
};

const StudentAttendanceApp: React.FC = () => {
  const [students, setStudents] = useState<Student[]>([]);
  const [uniqueClasses, setUniqueClasses] = useState<string[]>(["Semua"]);
  const [activeTab, setActiveTab] = useState<
    "attendance" | "recap" | "graph" | "semesterRecap"
  >("attendance");
  const [isAttendanceLoading, setIsAttendanceLoading] = useState(true);
  const [refreshTrigger, setRefreshTrigger] = useState(0);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [schoolData, setSchoolData] = useState<SchoolData | null>(null);
  const [studentsLoaded, setStudentsLoaded] = useState<boolean>(false);

  const fetchStudents = () => {
    fetch(endpoint)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data: Student[]) => {
        console.log("Data siswa yang diambil:", data);
        setStudents(data);

        const classSet = new Set<string>();
        data.forEach((student) => {
          if (student.kelas != null) {
            const kelasValue = String(student.kelas).trim();
            if (
              kelasValue !== "" &&
              kelasValue !== "undefined" &&
              kelasValue !== "null"
            ) {
              classSet.add(kelasValue);
            }
          }
        });
        const classes = Array.from(classSet).sort((a, b) => {
          const aIsNum = /^\d+$/.test(a);
          const bIsNum = /^\d+$/.test(b);
          if (aIsNum && bIsNum) return parseInt(a) - parseInt(b);
          if (aIsNum && !bIsNum) return -1;
          if (!aIsNum && bIsNum) return 1;
          return a.localeCompare(b);
        });
        setUniqueClasses(["Semua", ...classes]);
        setStudentsLoaded(true); // ✅ fetch selesai (berhasil)
      })
      .catch((error) => {
        console.error("Error fetch:", error);
        alert("❌ Gagal mengambil data siswa. Cek console untuk detail.");
        setStudentsLoaded(true); // ✅ fetch selesai (gagal, tapi tetap dianggap selesai)
      });
  };

  const fetchSchoolData = () => {
    fetch(`${endpoint}?action=schoolData`)
      .then((res) => {
        if (!res.ok) throw new Error(`HTTP error! status: ${res.status}`);
        return res.json();
      })
      .then((data) => {
        if (data.success && data.data && data.data.length > 0) {
          setSchoolData(data.data[0]);
          console.log("School data loaded:", data.data[0]);
        } else {
          setSchoolData(null);
        }
      })
      .catch((error) => {
        console.error("Error fetching school data:", error);
      });
  };

  const handleRecapRefresh = () => {
    setRefreshTrigger((prev) => prev + 1);
  };

  const handleRefresh = () => {
    fetchStudents();
    fetchSchoolData();
  };

  useEffect(() => {
    // Simulasi loading selama 3 detik
    const timer = setTimeout(() => {
      setIsLoading(false);
      fetchStudents();
      fetchSchoolData();
    }, 3000);

    return () => clearTimeout(timer);
  }, []);

  if (isLoading) {
    return <SplashScreen />;
  }

  const isGuruKelas = schoolData?.statusGuru === "Guru Kelas";
  const shouldShowJadwalMengajar = !isGuruKelas;

  return (
    <div className="min-h-screen bg-gray-100 flex flex-col">
      {/* Sidebar */}
      <aside
        className={`bg-white shadow-md w-64 space-y-2 py-6 px-2 fixed h-full top-0 left-0 transform ${
          isSidebarOpen ? "translate-x-0" : "-translate-x-full"
        } transition-transform duration-300 ease-in-out z-50`}
      >
        <div className="flex justify-between items-center mb-4 px-4">
          <h2 className="text-xl font-bold text-gray-800">Menu</h2>
          <button
            onClick={() => setIsSidebarOpen(false)}
            className="text-gray-600 hover:text-gray-800 text-2xl"
          >
            ✖️
          </button>
        </div>
        {[
          { tab: "attendance", label: "📋 Absensi" },
          { tab: "recap", label: "📊 Rekap Bulanan" },
          { tab: "semesterRecap", label: "📚 Rekap Semester" },
          { tab: "graph", label: "📈 Grafik" },
        ].map(({ tab, label }) => (
          <button
            key={tab}
            onClick={() => {
              if (tab === "attendance") {
                setIsAttendanceLoading(true);
              }
              setActiveTab(
                tab as "attendance" | "recap" | "graph" | "semesterRecap"
              );
              setIsSidebarOpen(false);
            }}
            className={`w-full text-left py-2 px-4 rounded-lg text-sm font-medium transition-colors flex items-center gap-2 ${
              activeTab === tab
                ? "bg-blue-600 text-white"
                : "text-gray-600 hover:bg-gray-100"
            }`}
          >
            {label}
          </button>
        ))}
      </aside>

      {/* Hamburger Menu Button */}
      <div className="fixed top-4 left-4 z-50">
        <button
          onClick={() => setIsSidebarOpen(!isSidebarOpen)}
          className="bg-blue-600 text-white px-4 py-2 rounded-lg shadow-md hover:bg-blue-700 transition-colors"
        >
          {isSidebarOpen ? "✖️ Tutup Menu" : "☰ Buka Menu"}
        </button>
      </div>

      {/* Logo di pojok kanan atas */}
      <div className="absolute top-4 right-4 z-50">
        <img
          src="\images\logo_2.png"
          alt="Logo Aplikasi"
          className="w-16 h-16"
        />
      </div>

      {/* Main Content */}
      <main
        className={`flex-1 p-6 transition-all duration-300 ${
          isSidebarOpen ? "ml-64" : "ml-0"
        } mt-16`}
      >
        <div className="text-center mb-6">
          <h1 className="text-3xl font-bold text-gray-800 mb-2">
            Sistem Absensi Siswa
          </h1>
          <p className="text-gray-600">Kelola data siswa dan absensi harian</p>
        </div>

        <div className="py-4">
          {activeTab === "attendance" && (
            <>
              {isAttendanceLoading && (
                <div className="flex flex-col items-center justify-center py-24">
                  <div className="bg-blue-50 border-2 border-blue-200 rounded-lg p-10 max-w-md mx-auto text-center">
                    <div className="text-6xl mb-4">⏳</div>
                    <h3 className="text-2xl font-bold text-blue-700 mb-2">
                      Mohon Tunggu
                    </h3>
                    <p className="text-blue-600 mb-4">
                      Sedang memuat data absensi...
                    </p>
                    <div className="flex justify-center">
                      <svg
                        className="animate-spin h-8 w-8 text-blue-600"
                        xmlns="http://www.w3.org/2000/svg"
                        fill="none"
                        viewBox="0 0 24 24"
                      >
                        <circle
                          className="opacity-25"
                          cx="12"
                          cy="12"
                          r="10"
                          stroke="currentColor"
                          strokeWidth="4"
                        ></circle>
                        <path
                          className="opacity-75"
                          fill="currentColor"
                          d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                        ></path>
                      </svg>
                    </div>
                  </div>
                </div>
              )}
              <div style={{ display: isAttendanceLoading ? "none" : "block" }}>
                <AttendanceTab
                  students={students}
                  onRecapRefresh={handleRecapRefresh}
                  onLoadingChange={setIsAttendanceLoading}
                  studentsLoaded={studentsLoaded}
                />
              </div>
            </>
          )}
          {activeTab === "recap" && (
            <MonthlyRecapTab
              onRefresh={handleRecapRefresh}
              uniqueClasses={uniqueClasses}
              students={students} // ✅ TAMBAHKAN INI
            />
          )}
          {activeTab === "graph" && <GraphTab uniqueClasses={uniqueClasses} />}
          {activeTab === "semesterRecap" && (
            <SemesterRecapTab
              uniqueClasses={uniqueClasses}
              students={students} // ✅ TAMBAHKAN INI
            />
          )}
        </div>
      </main>
    </div>
  );
};

const hasAnyAttendanceOnDate = (
  students: Student[],
  day: number,
  attendanceDataGetter: (student: Student) => {
    attendance: { [day: number]: string };
  }
): boolean => {
  return students.some((student) => {
    const { attendance } = attendanceDataGetter(student);
    return attendance[day] && attendance[day] !== "";
  });
};

export default StudentAttendanceApp;
