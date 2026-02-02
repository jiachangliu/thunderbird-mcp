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

# 2) Find the created draft by searching for TOK1.
search1=$(curl -sS -m 40 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"searchMessages","arguments":{"query":"'"$TOK1"'"}}}')

draft_id=$(python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); items=json.loads(j["result"]["content"][0]["text"]);
if len(items)!=1:
  print("FAIL: expected 1 search result for TOK1, got", len(items));
  [print(it.get("folderPath"), it.get("id")) for it in items];
  sys.exit(1)
print(items[0]["id"])' <<<"$search1")

draft_folder=$(python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); items=json.loads(j["result"]["content"][0]["text"]); print(items[0]["folderPath"])' <<<"$search1")

if [[ "$draft_folder" != "$DRAFTS" ]]; then
  echo "FAIL: expected draft in Drafts folder, got: $draft_folder" >&2
  exit 1
fi

# 3) Verify original draft contains token and quoted original snippet.
raw1=$(curl -sS -m 60 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"getRawMessage","arguments":{"messageId":"'"$draft_id"'","folderPath":"'"$DRAFTS"'"}}}')

python3 -c 'import json,re,sys; raw=sys.stdin.read(); raw=re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]","",raw); j=json.loads(raw); msg=json.loads(j["result"]["content"][0]["text"]); src=msg["source"]; body=src.split("\r\n\r\n",1)[1];
assert "'"$TOK1"'" in body, "missing TOK1";
# Quote should contain original email content:
assert "the committee very much enjoyed" in body, "missing quoted original snippet";
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
assert "the committee very much enjoyed" in body, "missing quoted original snippet after revise";
print("OK")' <<<"$raw2" >/dev/null

echo "PASS: reviseDraftInPlaceNativeEditor keeps quoted original"