import json, os

d = r'C:\Users\user\Desktop'
d = os.path.join(d, '\ub9c8\ub204\uc2a4 \uc790\uc7ac\uad00\ub9ac \uc5b4\ud50c\uac1c\ubc1c')

cfg = {
    "json_key_path": os.path.join(d, "primeval-span-492110-u0-f8910937cd53.json"),
    "spreadsheet_url": "https://docs.google.com/spreadsheets/d/1q0a6Vv7LYyIJSFcMwc2BMvRp8WUK-2TTTL_A1pYkcxI/edit?gid=0#gid=0"
}

config_path = os.path.join(d, "config.json")
with open(config_path, "w", encoding="utf-8") as f:
    json.dump(cfg, f, ensure_ascii=False, indent=2)

print("config.json updated successfully")
print("json_key_path:", cfg["json_key_path"])
print("exists:", os.path.exists(cfg["json_key_path"]))
