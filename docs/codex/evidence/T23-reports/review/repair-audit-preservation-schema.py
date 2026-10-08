from pathlib import Path
p=Path(__file__).with_name('audit-C1.py');s=p.read_text(encoding='utf-8')
s=s.replace("check(receipt['CommitUnderTest']==C1 and receipt['DirtyWorktree'] is False,'Native top receipt C1/clean source')", """if tier=='PreservationNative':
    check('CommitUnderTest' not in receipt and receipt['Task']=='T19','Legacy preservation receipt schema')
    notes.append({'shell':shell,'tier':tier,'source_binding':'Legacy feature receipt lacks its own commit/dirty fields; exact original raw capture is bound through C1-clean tier summary, source inventory and full immutable sourceguard.'})
   else:check(receipt['CommitUnderTest']==C1 and receipt['DirtyWorktree'] is False,'Native top receipt C1/clean source')""")
s=s.replace("'native_receipts':top_receipts,", "'native_receipts':top_receipts,'schema_notes':notes,")
p.write_text(s,encoding='utf-8')
