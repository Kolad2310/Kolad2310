```
import os

log_folder = "logs"
log_file = os.path.join(log_folder, "risk_engine.log")

# Create folder if not exists
os.makedirs(log_folder, exist_ok=True)

# Create empty log file
with open(log_file, "w") as f:
    pass

print("Empty log file created successfully.")
