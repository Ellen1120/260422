import json

data = json.loads(open('data/knowledge_base.json', encoding='utf-8').read())
kb = data['products'] if isinstance(data, dict) else data

# Rule 20: display_name 있는 항목
print('=== Rule 20: display_name (한글) ===')
found = False
for p in kb:
    for t in p.get('test_items', []):
        dn = t.get('display_name', '')
        if dn and dn != t.get('name', ''):
            print(f"  {p['stm_file']} | {t['name']} -> {dn}")
            found = True
if not found:
    print('  (없음)')

# Rule 19: 300287 이동상 재료
print()
print('=== Rule 19: STM-300287 이동상 ===')
p287 = next((p for p in kb if '300287' in p.get('stm_file', '')), None)
if p287:
    for ti in p287.get('test_items', []):
        for prep in ti.get('preparations', []):
            name = prep.get('solution_name', '')
            if 'mobile' in name.lower() or '이동상' in name:
                print(f'  prep: {name}')
                for ing in prep.get('ingredients', []):
                    print(f'    {ing}')
else:
    print('  300287 not found')

# CP 제품 test_items 확인
print()
print('=== CP 제품 test_items display_name 확인 ===')
cp_prods = [p for p in kb if p.get('stm_file', '').startswith('STM-CP')]
for p in cp_prods[:5]:
    tis = p.get('test_items', [])
    names = [(t.get('name'), t.get('display_name')) for t in tis[:5]]
    print(f"  {p['stm_file']}: {names}")
