#!/usr/bin/env bash
set -euo pipefail

# Cleanup script: remove test artifacts from Drafts.
# Deletes only messages that match known test patterns:
#  - Subject starting with "[TEST]"
#  - Draft ids starting with tb-mcp-draft-test-revise-
#  - Reply-draft artifacts for the integration tests, identified by body tokens TOK1-/TOK2-
#    (and subject "Re: Follow up" / "Follow up").

HOST="${HOST:-http://127.0.0.1:8765}"
DRAFTS="${DRAFTS:-imap://jl4624%40cornell.edu@outlook.office365.com/Drafts}"
LIMIT="${LIMIT:-250}"

python3 - <<'PY'
import json,os,re,subprocess,sys,time
HOST=os.environ.get('HOST','http://127.0.0.1:8765')
DRAFTS=os.environ.get('DRAFTS','imap://jl4624%40cornell.edu@outlook.office365.com/Drafts')
LIMIT=int(os.environ.get('LIMIT','250'))

TOK_RE=re.compile(r"\bTOK[0-9]?-\d+\b")

def rpc(name,args,timeout=60):
  payload={"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":name,"arguments":args}}
  out=subprocess.check_output(['curl','-sS','-m',str(timeout),'-X','POST',HOST,'-H','Content-Type: application/json','-d',json.dumps(payload)])
  j=json.loads(out)
  return json.loads(j['result']['content'][0]['text'])

def get_body(mid):
  obj=rpc('getRawMessage',{'messageId':mid,'folderPath':DRAFTS},timeout=60)
  src=obj.get('source','')
  if "\r\n\r\n" in src:
    return src.split("\r\n\r\n",1)[1]
  return src

lst=rpc('listLatestMessages',{'folderPath':DRAFTS,'limit':LIMIT},timeout=60)
items=lst.get('items',[])

to_delete=[]
for it in items:
  mid=it.get('id','')
  subj=(it.get('subject') or '')
  if not mid:
    continue

  # obvious test patterns
  if subj.startswith('[TEST]'):
    to_delete.append(mid)
    continue
  if mid.startswith('tb-mcp-draft-test-revise-'):
    to_delete.append(mid)
    continue

  # reply-draft artifacts: subject Follow up, body contains TOK*
  if subj.strip() in ('Re: Follow up','Follow up'):
    try:
      body=get_body(mid)
    except Exception:
      continue
    if TOK_RE.search(body):
      to_delete.append(mid)

# de-dupe preserving order
seen=set(); to_delete=[x for x in to_delete if not (x in seen or seen.add(x))]

if not to_delete:
  print('No matching test drafts found.')
  sys.exit(0)

print(f'Will delete {len(to_delete)} message(s) from Drafts.')

# Delete one-by-one (Draft deletions may be a move operation and can be slow).
for idx,mid in enumerate(to_delete, start=1):
  for attempt in range(1,4):
    try:
      res=rpc('deleteMessages',{'folderPath':DRAFTS,'messageIds':[mid]},timeout=300)
      if res.get('error'):
        raise RuntimeError(res['error'])
      print(f'Deleted {idx}/{len(to_delete)}: {mid}')
      break
    except Exception as e:
      if attempt >= 3:
        print('ERROR deleting',mid,'after retries:',e)
        sys.exit(2)
      print('Retry delete',mid,'attempt',attempt+1,'due to',e)
      time.sleep(2)
  time.sleep(0.2)

print('Cleanup complete.')
PY