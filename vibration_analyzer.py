import numpy as np
import pandas as pd

def calculate_metrics(signal_data):
    rms = np.sqrt(np.mean(np.square(signal_data)))
    peak = np.max(np.abs(signal_data))
    crest_factor = peak / rms if rms != 0 else 0
    
    return {
        "RMS": round(float(rms), 3),
        "Peak": round(float(peak), 3),
        "Crest_Factor": round(float(crest_factor), 3)
    }

def perform_fft(signal_data, sampling_rate=1000):
    n = len(signal_data)
    fft_vals = np.fft.rfft(signal_data)
    fft_freq = np.fft.rfftfreq(n, d=1.0 / sampling_rate)
    amplitude = np.abs(fft_vals) * 2.0 / n
    
    return pd.DataFrame({
        "Frequency_Hz": fft_freq,
        "Amplitude": amplitude
    })

def evaluate_status(rms_val, warning_threshold=2.5, danger_threshold=4.5):
    if rms_val < warning_threshold:
        return "🟢 정상", "success"
    elif rms_val < danger_threshold:
        return "🟡 주의", "warning"
    else:
        return "🔴 위험", "error"

def generate_site_sample_data(sampling_rate=1000):
    """
    여러 사업소(예: 부산사업소, 해운대사업소, 동래사업소 등)의 가상 진동 데이터 생성
    """
    t = np.linspace(0, 1, sampling_rate, endpoint=False)
    
    # 사업소 A (정상)
    sig_a = 1.2 * np.sin(2 * np.pi * 60 * t) + np.random.normal(0, 0.2, sampling_rate)
    
    # 사업소 B (주의 - 특정 주파수 튐)
    sig_b = 2.8 * np.sin(2 * np.pi * 60 * t) + 1.5 * np.sin(2 * np.pi * 120 * t) + np.random.normal(0, 0.5, sampling_rate)
    
    # 사업소 C (위험 - 진폭 큼)
    sig_c = 5.0 * np.sin(2 * np.pi * 60 * t) + 3.0 * np.sin(2 * np.pi * 180 * t) + np.random.normal(0, 0.8, sampling_rate)
    
    return {
        "부산 사업소": pd.DataFrame({"Time_s": t, "Vibration": sig_a}),
        "해운대 사업소": pd.DataFrame({"Time_s": t, "Vibration": sig_b}),
        "동래 사업소": pd.DataFrame({"Time_s": t, "Vibration": sig_c}),
    }