def asset_rows(rows):
    assert isinstance(rows, list) and len(rows) == 2
    expected = {'WinPDFMerger-v1.0.0.zip': (193669, ZIP), 'SHA256SUMS.txt': (90, SUMS)}
    assert len({x['name'] for x in rows}) == 2
    assert {x['name']: (x['bytes'], x['sha256']) for x in rows} == expected


def validate_host(report, kind, child, native):
    count, invocations, version, edition = {'PS51': (14, 54, '5.1.26100.9444', 'Desktop'),
                                           'PS7': (11, 51, '7.6.6', 'Core')}[kind]
    assert report['task'] == 'T33' and report['result'] == 'pass' and report['preparation'] is False
    assert report['shell_kind'] == kind and len(report['cases']) == count
    assert report['evidence_class'] == CLASS and report['harness_commit'] == M and report['candidate_source_commit'] == R
    guard_keys = {'expected_head', 'status_unchanged', 'clean', 'driver_unchanged',
                  'approved_cache_unchanged', 'candidate_assets_unchanged', 'parent_environment_unchanged'}
    assert set(report['source_guard']) == guard_keys and all(report['source_guard'][x] is True for x in guard_keys)
    assert report['manual_acceptance'] == 'excluded/unperformed; never pass'
    assert all(x['package_guard'] is True and x['source_foreign_guard'] is True for x in report['cases'])
    assert child['shell'] == kind and child['cases'] == count and child['invocations'] == invocations and child['result'] == 'pass'
    assert report['invocation_count'] == invocations
    assert len(report['approved_cache_files']) == native['approved_cache_files_verified'] == 348
    assert report['driver_sha256_before'] == report['driver_sha256_after'] == native['harness_sha256']
    for key in ('zip_path', 'checksums_path'):
        assert Path(report['candidate'][key]).resolve() == Path(native['shared_assets'][key]).resolve()
    assert report['candidate']['zip_sha256'] == ZIP and report['candidate']['checksums_sha256'] == SUMS
    assert report['candidate']['build_info']['source_commit'] == R
    env = report['environment']
    assert env['shell_version'] == version and env['shell_edition'] == edition
    assert env['process_64_bit'] is True and env['is_administrator'] is False
    assert (env['edition'], env['display_version'], env['full_build']) == ('Professional', '26H2', '26300.9457')
    policies = {x['scope']: x['policy'] for x in env['policy']}
    assert policies['Process'] == 'RemoteSigned' and policies['MachinePolicy'] == policies['UserPolicy'] == 'Undefined'


def validate_bindings(g, paths, repo):
    download, package, native, operation, images, visual = (g[x] for x in
        ('public_download', 'package_review', 'actual_operation', 'operation_review', 'decoded_image_review', 'root_visual_review'))

    def linked(row, path):
        assert (repo / row['path']).resolve() == path.resolve()
        assert row['sha256'] == sha(path)

    linked(download['package_audit'], paths['package_review'])
    assert download['package_audit']['checks'] == package['checks_total'] and download['package_audit']['issues'] == []
    directory = Path(download['download_directory']).resolve()
    assert directory.is_dir() and not directory.is_relative_to(repo)
    assert Path(operation['download_directory']).resolve() == directory
    for key, name, digest in [('zip_path', 'WinPDFMerger-v1.0.0.zip', ZIP), ('checksums_path', 'SHA256SUMS.txt', SUMS)]:
        assert Path(native['shared_assets'][key]).resolve() == directory / name
        assert sha(directory / name) == digest
    linked(operation['public_download_report'], paths['public_download'])
    linked(operation['actual_native_ledger'], paths['actual_operation'])
    linked(operation['decoded_image_review'], paths['decoded_image_review'])
    original = (repo / operation['original_operation_report']['path']).resolve()
    execution = (repo / operation['actual_execution_receipt']['path']).resolve()
    assert original.is_relative_to(repo / 'tests/.work') and execution.is_relative_to(repo / 'tests/.work')
    linked(operation['original_operation_report'], original)
    linked(operation['actual_execution_receipt'], execution)
    assert operation['actual_execution_receipt']['exit_code'] == read(execution)['exit_code'] == 0
    original_report = read(original)
    assert original_report['task'] == 'T33' and original_report['issues'] == []
    assert original_report['candidate_source_commit'] == R and original_report['harness_commit'] == M
    assert original_report['application_reexecuted'] is False and original_report['pdftk_or_ghostscript_reexecuted'] is False
    assert original_report['checks'] == operation['checks']
    assert len(original_report['actual_application_cases']) == 25 and len(original_report['independent_pdf_reads']) == 21
    assert sum(len(x['pages']) for x in original_report['independent_pdf_reads']) == 106
    children = native['candidate_reports']
    assert isinstance(children, list) and len(children) == 2 and {x['shell'] for x in children} == {'PS51', 'PS7'}
    receipts, host_reports = [], {}
    for kind in ('PS51', 'PS7'):
        child = next(x for x in children if x['shell'] == kind)
        path = paths['actual_operation'].parent / (kind + '-reports') / 'result.json'
        linked(child, path)
        report = read(path)
        validate_host(report, kind, child, native)
        host_reports[kind] = (path, report)
        receipts.append({'path': kind + '-reports/result.json', 'sha256': sha(path)})
    assert images['input_receipts'] == receipts

    render_path = paths['decoded_image_review'].parent / 'visual-published-outputs/render-contact-report.json'
    independent_path = paths['decoded_image_review'].parent / 'visual-published-review.json'
    render, independent = read(render_path), read(independent_path)
    assert render['task'] == 'T33' and render['source_commit'] == R
    assert render['result'] == 'rendered_pending_visual_review'
    assert Path(render['capture']).resolve() == paths['actual_operation'].parent.resolve()
    assert (render['pdf_count'], render['page_count'], len(render['sheets'])) == (21, 106, 6)
    assert independent['task'] == 'T33' and independent['result'] == 'pass_for_recorded_rendered_contact_sheet_inspection'
    assert independent['issues'] == [] and independent['source_commit'] == R and independent['harness_commit'] == M
    assert independent['zip_sha256'] == ZIP and independent['checksums_sha256'] == SUMS
    assert independent['original_render_report_sha256'] == sha(render_path)
    assert (independent['pdf_count'], independent['page_count'], independent['contact_sheets'], len(independent['sheets'])) == (21, 106, 6, 6)
    assert independent['human_acceptance'] == 'excluded/unperformed; never pass'
    assert independent['application_native_or_download_reexecuted'] is False
    assert len({x['path'] for x in render['sheets']}) == 6
    for sheet in render['sheets']:
        path = (render_path.parent / sheet['path']).resolve()
        assert path.is_relative_to(render_path.parent) and sha(path) == sheet['sha256']
        matches = [x for x in independent['sheets'] if (independent_path.parent / x['path']).resolve() == path]
        assert len(matches) == 1 and matches[0]['sha256'] == sheet['sha256'] and matches[0]['bytes'] == path.stat().st_size
    expected_root_sheets = ['contact-01.png', 'contact-04.png']
    for row, name, prefix in zip(visual['sheets'], expected_root_sheets, ('PS51-', 'PS7-')):
        assert row['path'].startswith('<REPO>/')
        path = (repo / row['path'][len('<REPO>/'):]).resolve()
        assert path == (render_path.parent / name).resolve() and row['sha256'] == sha(path)
        sheet = next(x for x in render['sheets'] if x['path'] == name)
        assert row['sha256'] == sheet['sha256']
        assert sheet['pdf_labels'] == [prefix + x for x in ('default-screen-master', 'default-screen-email', 'ebook-output-master', 'ebook-output-email')]
    assert visual['tool'] == 'view_image'
    return host_reports, render_path, independent_path
