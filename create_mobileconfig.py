"""
Generate dashboard.mobileconfig — iOS Web Clip profile.
Install on iPhone: open the file (via AirDrop, email, Telegram, or http://PC_IP:8080/install)
Result: Financial Dashboard icon appears on home screen, opens in full-screen Safari.
"""
import base64, os, uuid

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ICON_PATH  = os.path.join(SCRIPT_DIR, "apple-touch-icon.png")
OUT_PATH   = os.path.join(SCRIPT_DIR, "dashboard.mobileconfig")

URL         = "https://muktihadi5641-cpu.github.io/financial-bot/"
LABEL       = "Financial"
PROFILE_ID  = "com.mukti.financialdashboard"
CLIP_ID     = "com.mukti.financialdashboard.webclip"
PROFILE_UUID = str(uuid.uuid4()).upper()
CLIP_UUID    = str(uuid.uuid4()).upper()

with open(ICON_PATH, "rb") as f:
    icon_b64 = base64.b64encode(f.read()).decode()

plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
\t<key>PayloadContent</key>
\t<array>
\t\t<dict>
\t\t\t<key>PayloadType</key>
\t\t\t<string>com.apple.webClip.managed</string>
\t\t\t<key>PayloadVersion</key>
\t\t\t<integer>1</integer>
\t\t\t<key>PayloadIdentifier</key>
\t\t\t<string>{CLIP_ID}</string>
\t\t\t<key>PayloadUUID</key>
\t\t\t<string>{CLIP_UUID}</string>
\t\t\t<key>Label</key>
\t\t\t<string>{LABEL}</string>
\t\t\t<key>URL</key>
\t\t\t<string>{URL}</string>
\t\t\t<key>IsRemovable</key>
\t\t\t<true/>
\t\t\t<key>FullScreen</key>
\t\t\t<true/>
\t\t\t<key>IgnoreManifestScope</key>
\t\t\t<false/>
\t\t\t<key>Icon</key>
\t\t\t<data>{icon_b64}</data>
\t\t</dict>
\t</array>
\t<key>PayloadDisplayName</key>
\t<string>Financial Dashboard</string>
\t<key>PayloadIdentifier</key>
\t<string>{PROFILE_ID}</string>
\t<key>PayloadRemovalDisallowed</key>
\t<false/>
\t<key>PayloadType</key>
\t<string>Configuration Profile</string>
\t<key>PayloadUUID</key>
\t<string>{PROFILE_UUID}</string>
\t<key>PayloadVersion</key>
\t<integer>1</integer>
</dict>
</plist>"""

with open(OUT_PATH, "w", encoding="utf-8") as f:
    f.write(plist)

kb = os.path.getsize(OUT_PATH) / 1024
print(f"dashboard.mobileconfig created ({kb:.1f} KB)")
print(f"URL: {URL}")
