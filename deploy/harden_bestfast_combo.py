#!/usr/bin/env python3
"""Harden OmniRoute auto/best-fast — kill combo burners, keep proven winners.

2026-10-06: OWUI same-chat failed with "All credentials cooling down" after
Excel FULL_FILE_READ succeeded. Root cause = virtual auto hunt burned through
zai/huggingface/groq/mistral/chatgpt-web/… before gemini/alibaba could win.

KEEP ON (proven 200 today / tools-capable):
  gemini, alibaba, cloudflare-ai, opencode, deepseek, grok-cli

FORCE OFF (burn retries / 403/400/429/no-balance):
  zai, github, huggingface, groq, mistral, amazon-q, agentrouter, cline,
  chatgpt-web, claude-web, jules, cerebras, devin
  (+ prior dead set already off)
"""
from __future__ import annotations

import sys

import paramiko

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

HOST = "192.168.0.17"

REMOTE = r'''
import json, os, shutil, sqlite3, time, urllib.request, re, subprocess
from datetime import datetime, timezone

OMNI = "/opt/omniroute/data/storage.sqlite"

# provider -> id (verified live)
KEEP = {
    "dd6e6fde-7218-4d5e-9a49-ccf3e7c41e9d": "gemini",
    "e911c71a-1361-442b-89d0-0de10316665c": "alibaba",
    "81f7f38d-9696-4521-8188-117aeadf2c79": "cloudflare-ai",
    "a74a2795-6a4f-4945-b454-9512215147d5": "opencode",
    "65f11544-aa5b-4a68-927b-fd20e7005782": "deepseek",
    "e413867d-c5a1-41a2-812b-ff606b94ab5f": "grok-cli",
}

OFF = {
    "3b284922-4835-4d47-9ec2-64cdf7a54a87": "zai",
    "8067d69e-256a-4c47-90db-c25b156d8974": "github",
    "de252c27-1722-46a3-81be-4278f150b033": "huggingface",
    "90b126c0-6f83-4cbc-8f53-568b65f0747e": "groq",
    "bacf8e1b-4672-433f-baaa-2273946b1330": "mistral",
    "545b1f59-dc65-4f55-b4d6-3efc506e1520": "amazon-q",
    "0e6044b7-70cf-4fb9-acd3-79daede794e1": "agentrouter",
    "b2a7e7aa-c268-4fa1-b0d4-afab17f8c2ff": "cline",
    "8ddc3ce7-0ef8-4456-a60a-21f4d4287206": "chatgpt-web",
    "75456dd3-6171-44a7-bce7-c0869386eb98": "claude-web",
    "69925c5e-1e8a-4c6d-8aed-2ad785c723ff": "jules",
    "c44eeda1-5069-46a1-9bdd-3bd9055c7696": "cerebras",
    "ec689cb8-ceb1-4562-9613-4a479f955016": "devin",
    # prior dead set (idempotent)
    "41bd143b-f816-478f-b2c6-b8d647bf1405": "opencode-go",
    "7a08c5d5-8b6b-4b4f-8618-4019365f7a42": "kiro",
    "3450a44a-caba-4cf8-971c-82ef96ab82fa": "kilocode",
    "d07df148-de07-445c-9e90-b648c3fcae1c": "devin-cli",
    "2aa9214d-0b30-4f88-8a6e-7860bc868cac": "gemini-web",
    "1a30a84f-a8f8-4bd5-b209-b3b3b037be2a": "zenmux",
    "33abc315-b3c9-4260-b343-5df20950cea8": "deepinfra",
    "4d84a13b-a3c1-4936-9281-d4a2ecb0a845": "agnes",
    "f0419e9e-8f98-4971-81b2-08d6ae9c28b4": "deepseek-web",
}

ts = datetime.now().strftime("%Y%m%d-%H%M%S")
bak = f"{OMNI}.bak-combo-harden-{ts}"
os.system(f'sqlite3 "{OMNI}" "PRAGMA wal_checkpoint(TRUNCATE);" >/dev/null 2>&1')
shutil.copy2(OMNI, bak)
os.chmod(bak, 0o600)
print("BACKUP", bak)

con = sqlite3.connect(OMNI)
con.row_factory = sqlite3.Row
now = datetime.now(timezone.utc).isoformat()

print("=== OFF ===")
for cid, name in OFF.items():
    r = con.execute("select provider, is_active from provider_connections where id=?", (cid,)).fetchone()
    if not r:
        print("MISSING", name, cid)
        continue
    if r["provider"] != name:
        raise SystemExit(f"ID mismatch expected {name} got {r['provider']}")
    con.execute(
        "update provider_connections set is_active=0, updated_at=? where id=? and provider=?",
        (now, cid, name),
    )
    print(f"  {name}: {r['is_active']} -> 0")

print("=== KEEP + clear backoff ===")
for cid, name in KEEP.items():
    r = con.execute("select provider, is_active, rate_limited_until, backoff_level from provider_connections where id=?", (cid,)).fetchone()
    if not r:
        print("MISSING_KEEP", name)
        continue
    if r["provider"] != name:
        raise SystemExit(f"KEEP mismatch {name}")
    con.execute(
        """update provider_connections set is_active=1, rate_limited_until=NULL,
           backoff_level=0, last_error=NULL, error_code=NULL, updated_at=?
           where id=? and provider=?""",
        (now, cid, name),
    )
    # Prefer gemini/alibaba/opencode in account priority (lower = earlier if used)
    prio = {"gemini": 0, "alibaba": 0, "opencode": 0, "cloudflare-ai": 1, "deepseek": 1, "grok-cli": 1}[name]
    con.execute(
        "update provider_connections set priority=?, global_priority=? where id=?",
        (prio, prio, cid),
    )
    print(f"  {name}: active=1 prio={prio} (was active={r['is_active']} backoff={r['backoff_level']})")

# Soften model lockout so one 429 does not brick same chat for 2+ minutes
# Rewrite whole object (avoid duplicate/wrong key names)
row = con.execute("select value from key_value where namespace='settings' and key='modelLockout'").fetchone()
print("modelLockout BEFORE_RAW", row[0] if row else None)
ml = {
    "enabled": True,
    "errorCodes": [403, 404, 429, 502, 503, 504],
    "baseCooldownMs": 15000,   # was 120000
    "maxCooldownMs": 120000,    # was 1800000
    "maxBackoffSteps": 4,
    "useExponentialBackoff": True,
}
con.execute(
    "update key_value set value=? where namespace='settings' and key='modelLockout'",
    (json.dumps(ml, separators=(",", ":")),),
)
print("modelLockout AFTER", ml)

# bump settings revision so caches reload
con.execute(
    "update key_value set value=? where namespace='settings' and key='_settingsRevision'",
    (str(int(time.time())),),
)

con.commit()

print("\n=== ACTIVE NOW ===")
for r in con.execute("select provider, is_active, priority, global_priority from provider_connections order by provider"):
    mark = "ON " if r["is_active"] else "off"
    print(mark, r["provider"], "prio", r["priority"], r["global_priority"])

key = con.execute("select key from api_keys where is_active=1 and name='ghazala_key'").fetchone()[0]
con.close()

print("\n=== restart omniroute ===")
os.system("docker restart omniroute")
for i in range(20):
    time.sleep(3)
    code = os.popen("curl -sS -o /dev/null -w '%{http_code}' http://192.168.0.17:20128/ 2>/dev/null").read().strip()
    print(f"wait {i}: {code}")
    if code == "200":
        break

def smoke(label, body, timeout=180):
    req = urllib.request.Request(
        "http://192.168.0.17:20128/v1/chat/completions",
        data=json.dumps(body).encode(),
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        method="POST",
    )
    t0 = time.time()
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode()
            http = resp.status
    except Exception as ex:
        print(label, "FAIL", type(ex).__name__, str(ex)[:400], "sec", round(time.time()-t0,1))
        return False
    d = json.loads(raw)
    sec = round(time.time()-t0, 1)
    if d.get("error"):
        print(label, "ERROR", str(d["error"])[:400], "sec", sec)
        return False
    ch = (d.get("choices") or [{}])[0]
    msg = (ch.get("message") or {})
    content = msg.get("content")
    tool_calls = msg.get("tool_calls")
    print(label, "HTTP", http, "sec", sec, "MODEL", d.get("model"),
          "CONTENT", repr(content)[:80] if content else None,
          "TOOLS", bool(tool_calls))
    return http == 200

print("\n=== SMOKE plain ===")
ok1 = smoke("plain", {
    "model": "auto/best-fast",
    "messages": [{"role": "user", "content": "Reply with exactly: OK"}],
    "max_tokens": 16,
    "stream": False,
})

print("\n=== SMOKE tools (OWUI-like) ===")
ok2 = smoke("tools", {
    "model": "auto/best-fast",
    "messages": [{"role": "user", "content": "Say hi in one word. Do not call tools."}],
    "max_tokens": 32,
    "stream": False,
    "tools": [{
        "type": "function",
        "function": {
            "name": "generate_excel",
            "description": "Create Excel workbook",
            "parameters": {"type": "object", "properties": {"content": {"type": "string"}}, "required": []},
        },
    }, {
        "type": "function",
        "function": {
            "name": "generate_forecast_pack",
            "description": "Create forecast pack",
            "parameters": {"type": "object", "properties": {"job": {"type": "string"}}, "required": []},
        },
    }],
    "tool_choice": "none",
})

print("\n=== recent log ===")
out = subprocess.check_output(
    "grep -aE 'Auto selection:|succeeded|All models failed|Maximum combo|Virtual auto-combo' "
    "/opt/omniroute/data/logs/application/app.log | tail -30",
    shell=True, text=True,
)
for ln in out.splitlines():
    m = re.search(r'"msg":"([^"]+)"', ln)
    t = re.search(r'"time":"([^"]+)"', ln)
    if m:
        print((t.group(1) if t else "")[:19], m.group(1)[:200])

print("\nSMOKE_PLAIN", ok1, "SMOKE_TOOLS", ok2)
print("DONE", "PASS" if (ok1 and ok2) else "PARTIAL")
'''

c = paramiko.SSHClient()
c.set_missing_host_key_policy(paramiko.AutoAddPolicy())
c.connect(HOST, username="root", timeout=30)
sftp = c.open_sftp()
with sftp.file("/tmp/harden_bestfast_combo.py", "w") as f:
    f.write(REMOTE)
sftp.close()
_, o, e = c.exec_command("python3 /tmp/harden_bestfast_combo.py", timeout=360)
print((o.read() + e.read()).decode("utf-8", "replace"))
c.close()
