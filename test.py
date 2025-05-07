import psutil

print("目前執行中的 process 名稱：")
for proc in psutil.process_iter(attrs=["pid", "name"]):
    try:
        print(f"PID {proc.info['pid']}: {proc.info['name']}")
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass  # 有些系統 process 可能無法存取，略過
