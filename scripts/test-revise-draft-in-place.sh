#!/usr/bin/env bash
set -euo pipefail

# Integration test against a running Thunderbird MCP server on :8765
# Creates a unique draft, revises it in-place, and asserts:
#  - resulting draft contains new content
#  - there is only one draft with the subject (no duplicates)

HOST="http://localhost:8765"
FOLDER="imap://jl4624%40cornell.edu@outlook.office365.com/Drafts"
SUBJECT="[TEST] reviseDraftInPlaceNativeEditor $(date +%s)"

create_json=$(cat <<JSON
{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"saveDraft","arguments":{"to":"jl4624@cornell.edu","subject":"$SUBJECT","body":"Old body line 1\n\nBest,\nJiachang","isHtml":false,"idempotencyKey":"test-revise-$(date +%s)"}}}
JSON
)

resp=$(curl -sS -m 30 -X POST "$HOST" -H 'Content-Type: application/json' -d "$create_json")
echo "$resp" | head -c 400 >/dev/null

# Find the created draft id by listing Drafts.
list=$(curl -sS -m 30 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"listLatestMessages","arguments":{"folderPath":"'"$FOLDER"'","limit":20}}}')

draft_id=$(python3 - <<PY
import json,re,sys
raw=sys.stdin.read()
# strip control chars
raw=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', raw)
j=json.loads(raw)
text=j['result']['content'][0]['text']
text=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
obj=json.loads(text)
subj="$SUBJECT"
for it in obj.get('items', []):
    if it.get('subject')==subj:
        print(it.get('id'))
        sys.exit(0)
print('')
sys.exit(1)
PY
<<<"$list")

revise=$(curl -sS -m 120 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"reviseDraftInPlaceNativeEditor","arguments":{"messageId":"'"$draft_id"'","folderPath":"'"$FOLDER"'","plainTextBody":"New body line 1\nNew body line 2\n\nBest,\nJiachang","closeAfterSave":true}}}')

echo "$revise" | head -c 400 >/dev/null

# Re-list and ensure only one matching subject.
list2=$(curl -sS -m 30 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"listLatestMessages","arguments":{"folderPath":"'"$FOLDER"'","limit":40}}}')

python3 - <<PY
import json,re,sys
subj="$SUBJECT"
raw=sys.stdin.read()
raw=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', raw)
j=json.loads(raw)
text=j['result']['content'][0]['text']
text=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
obj=json.loads(text)
items=[it for it in obj.get('items',[]) if it.get('subject')==subj]
if len(items)!=1:
    print('FAIL: expected 1 draft with subject, got', len(items))
    for it in items:
        print(it.get('id'), it.get('date'))
    sys.exit(1)
print(items[0]['id'])
PY
<<<"$list2" > /tmp/test_revise_new_id.txt

new_id=$(cat /tmp/test_revise_new_id.txt)
rawmsg=$(curl -sS -m 60 -X POST "$HOST" -H 'Content-Type: application/json' -d '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"getRawMessage","arguments":{"messageId":"'"$new_id"'","folderPath":"'"$FOLDER"'"}}}')

python3 - <<PY
import json,re,sys
raw=sys.stdin.read()
raw=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', raw)
j=json.loads(raw)
text=j['result']['content'][0]['text']
text=re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)
msg=json.loads(text)
src=msg['source']
body=src.split('\r\n\r\n',1)[1]
if 'New body line 1' not in body:
    print('FAIL: revised body not found')
    sys.exit(1)
print('OK')
PY
<<<"$rawmsg"

echo "PASS: reviseDraftInPlaceNativeEditor"