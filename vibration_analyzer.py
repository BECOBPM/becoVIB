import numpy as np
import pandas as pd

def calculate_metrics(signal_data):
    """
    진동 데이터의 주요 수치 지표 계산 (RMS, Peak, Crest Factor)
    """
    rms = np.sqrt(np.mean(np.square(signal_data)))
    peak = np.max(np.abs(signal_data))
    crest_factor = peak / rms if rms != 0 else 0
    
    return {
        "RMS": round(float(rms), 3),
        "Peak": round(float(peak), 3),
        "Crest_Factor": round(float(crest_factor), 3)
    }

def perform_fft(signal_data, sampling_rate=1000):
    """
    고속 푸리에 변환(FFT)을 통한 주파수 스펙트럼 데이터 생성
    """
    n = len(signal_data)
    fft_vals = np.fft.rfft(signal_data)
    fft_freq = np.fft.rfftfreq(n, d=1.0 / sampling_rate)
    amplitude = np.abs(fft_vals) * 2.0 / n
    
    return pd.DataFrame({
        "Frequency_Hz": fft_freq,
        "Amplitude": amplitude
    })

def evaluate_status(rms_val, warning_threshold=2.8, danger_threshold=4.5):
    """
    RMS 진동값 기준 상태 판정 (ISO 기준 참고)
    """
    if rms_val < warning_threshold:
        return "🟢 정상 (Normal)", "success"
    elif rms_val < danger_threshold:
        return "🟡 주의 (Warning)", "warning"
    else:
        return "🔴 위험 (Danger)", "error"