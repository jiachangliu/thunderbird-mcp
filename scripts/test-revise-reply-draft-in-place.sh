#!/usr/bin/env bash
set -euo pipefail

# Integration test: create a REPLY draft with quoted original, then revise it in-place.
# Asserts:
#  - revised draft contains new top text
#  - quoted original content is still present after revision

HOST="http://localhost:8765"
INBOX="imap://jl4624%40cornell.edu@outlook.office365.com/INBOX"
DRAFTS="imap://jl4624%40cornell.edu@outlook.office365.com/Drafts"

# Use Becky Passonneau "Follow up" email as a stable source message.
SRC_MSGID="eac45bf2-9d0a-4af9-92b7-0514558ee366@psu.edu"

TOK1="TOK1-$(date +%s)-$RANDOM"
TOK2="TOK2-$(date +%s)-$RANDOM"

# 1) Create a reply-all draft with quoted original.
create=$(curl -sS -m 180 -X POST "$HOST" -H 'Content-Type: application/json' -d '{
  "jsonrpc":"2.0",
  "id":1,
  "method":"tools/call",
  "params":{
    "name":"replyToMessageDraftNativeEditor",
    "arguments":{
      "messageId":"'"$SRC_MSGID"'",
      "folderPath":"'"$INBOX"'",
      "replyAll":true,
      "plainTextBody":"Hi Becky,\n\n'$TOK1'\n\nBest,\nJiachang",
      "closeAfterSave":true
    }
  }
}')

echo "$create" | head -c 200 >/dev/null

# 2) Find the created draft by polling Drafts and searching raw MIME for TOK1.
# (Thunderbird search index may lag, so don't rely on searchMessages here.)

draft_id=""
for i in {1..30}; do
  list=$(curl -sS -m 30 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"listLatestMessages","arguments":{"folderPath":"'"$DRAFTS"'","limit":20}}}')
  # Extract ids from list and scan raw for TOK1.
  ids=$(python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); text=j["result"]["content"][0]["text"]; text=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",text); obj=json.loads(text); print("\n".join([it.get("id","") for it in obj.get("items",[])]))' <<<"$list")
  while IFS= read -r id; do
    [[ -z "$id" ]] && continue
    raw=$(curl -sS -m 30 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"getRawMessage","arguments":{"messageId":"'"$id"'","folderPath":"'"$DRAFTS"'"}}}')
    found=$(python3 -c 'import json,re,sys; tok=sys.argv[1]; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); msg=json.loads(j["result"]["content"][0]["text"]); src=msg["source"]; body=src.split("\r\n\r\n",1)[1]; print("YES" if tok in body else "NO")' "$TOK1" <<<"$raw")
    if [[ "$found" == "YES" ]]; then
      draft_id="$id"
      raw1="$raw"
      break
    fi
  done <<<"$ids"

  if [[ -n "$draft_id" ]]; then
    break
  fi
  sleep 2
done

if [[ -z "$draft_id" ]]; then
  echo "FAIL: could not find reply draft containing TOK1 in Drafts" >&2
  exit 1
fi

# 3) Verify original draft contains token and quoted original snippet.

python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); msg=json.loads(j["result"]["content"][0]["text"]); src=msg["source"]; body=src.split("\r\n\r\n",1)[1];
assert "'"$TOK1"'" in body, "missing TOK1";
# Quote should contain original email content:
assert "thank you again for your time last week" in body.lower(), "missing quoted original snippet";
print("OK")' <<<"$raw1" >/dev/null

# 4) Revise in-place (replace body). Keep quote expected to remain.
rev=$(curl -sS -m 240 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"reviseDraftInPlaceNativeEditor","arguments":{"messageId":"'"$draft_id"'","folderPath":"'"$DRAFTS"'","plainTextBody":"Hi Becky,\n\n'$TOK2'\n\nBest,\nJiachang","closeAfterSave":true}}}')

echo "$rev" | head -c 200 >/dev/null

new_id=$(python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); obj=json.loads(j["result"]["content"][0]["text"]); 
if not obj.get("ok"):
  print("FAIL:", obj); sys.exit(1)
print(obj.get("messageId"))' <<<"$rev")

# 5) Verify revised draft contains TOK2 and still contains quoted original snippet.
raw2=$(curl -sS -m 60 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"getRawMessage","arguments":{"messageId":"'"$new_id"'","folderPath":"'"$DRAFTS"'"}}}')

python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); msg=json.loads(j["result"]["content"][0]["text"]); src=msg["source"]; body=src.split("\r\n\r\n",1)[1];
assert "'"$TOK2"'" in body, "missing TOK2";
assert "thank you again for your time last week" in body.lower(), "missing quoted original snippet after revise";
print("OK")' <<<"$raw2" >/dev/null

echo "PASS: reviseDraftInPlaceNativeEditor keeps quoted original"