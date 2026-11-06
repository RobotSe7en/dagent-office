import { readFileSync } from 'node:fs';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createAssetDescriptor, createWorkbook, openDocument } from '@doffi/core';
import type { WorkbookManifest } from '@doffi/core';
import { modelAssetSource, prepareWorkbookAssets } from '../src/workbook/assets.js';
import type { WorkbookAssetResponse, WorkbookAssetSource } from '../src/workbook/assets.js';
const bytes = new Uint8Array(readFileSync('examples/presentation/assets/sample.png'));
function seed() {
  const files = new Map(createWorkbook({ documentId: 'display-assets' })), manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as WorkbookManifest;
  for (const path of ['assets/one.png', 'assets/two.png']) { manifest.assets.push(createAssetDescriptor(path, 'image/png', bytes)); files.set(path, bytes); }
  files.set('manifest.json', new TextEncoder().encode(JSON.stringify(manifest)));
  const model = openDocument(files, { kind: 'workbook' });
  return modelAssetSource({ ...model, serialize() { throw Error('Display must not serialize the workbook'); } });
}
const signal = () => new AbortController().signal;
afterEach(() => vi.restoreAllMocks());

describe('workbook window display resources', () => {
  it('deduplicates one window, isolates independent leases and releases URLs idempotently', async () => {
    const source = seed(), create = vi.spyOn(URL, 'createObjectURL'), revoke = vi.spyOn(URL, 'revokeObjectURL');
    const first = await prepareWorkbookAssets(source, ['assets/one.png', 'assets/one.png'], signal());
    const second = await prepareWorkbookAssets(source, ['assets/one.png'], signal());
    try {
      expect(first.count).toBe(1); expect(first.byteLength).toBe(bytes.length); expect(first.pixels).toBe(64); expect(first.url('unknown')).toBeUndefined();
      expect(first.url('assets/one.png')).not.toBe(second.url('assets/one.png')); expect(create).toHaveBeenCalledTimes(2);
      expect(new Uint8Array(await (await fetch(first.url('assets/one.png')!)).arrayBuffer())).toEqual(bytes);
      expect(Object.isFrozen(first)).toBe(true); first.dispose(); first.dispose(); expect(revoke).toHaveBeenCalledTimes(1);
      expect(() => first.url('assets/one.png')).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
      expect((await fetch(second.url('assets/one.png')!)).ok).toBe(true);
    } finally { first.dispose(); second.dispose(); }
    expect(revoke).toHaveBeenCalledTimes(2);
  });
  it('validates all requested paths and count limits before contacting the provider', async () => {
    const source = seed(), readAsset = vi.fn(source.readAsset), guarded = { ...source, readAsset };
    await expect(prepareWorkbookAssets(guarded, ['assets/one.png', 'styles.css'], signal())).rejects.toMatchObject({ code: 'AssetUnavailable' });
    await expect(prepareWorkbookAssets(guarded, ['assets/one.png', 'assets/two.png'], signal(), { maxAssets: 1 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    await expect(prepareWorkbookAssets(guarded, Array(3).fill('assets/one.png'), signal(), { maxPaths: 2 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    for (const maxBytes of [0, -1, 1.5, Infinity]) await expect(prepareWorkbookAssets(guarded, [], signal(), { maxBytes })).rejects.toMatchObject({ code: 'InvalidOperation' });
    expect(readAsset).not.toHaveBeenCalled();
    const snapshot = { ...source.snapshot, manifest: { ...source.snapshot.manifest, assets: source.snapshot.manifest.assets.map(asset => ({ ...asset, mimeType: 'image/svg+xml' })) } };
    await expect(prepareWorkbookAssets({ ...guarded, snapshot }, ['assets/one.png'], signal())).rejects.toMatchObject({ code: 'UnsupportedFeature' });
    expect(readAsset).not.toHaveBeenCalled();
  });
  it('passes the remaining byte budget to the model and releases earlier resources on failure', async () => {
    const source = seed(), readAsset = vi.fn(source.readAsset), create = vi.spyOn(URL, 'createObjectURL'), revoke = vi.spyOn(URL, 'revokeObjectURL');
    await expect(prepareWorkbookAssets({ ...source, readAsset }, ['assets/one.png', 'assets/two.png'], signal(), { maxAssetBytes: bytes.length, maxBytes: bytes.length * 2 - 1 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(readAsset.mock.calls.map(([request]) => request.maxBytes)).toEqual([bytes.length, bytes.length - 1]);
    expect(create).toHaveBeenCalledTimes(1); expect(revoke).toHaveBeenCalledWith(create.mock.results[0]!.value);
  });
  it('checks individual and cumulative pixels before making URLs and releases earlier resources on failure', async () => {
    const source = seed(), create = vi.spyOn(URL, 'createObjectURL'), revoke = vi.spyOn(URL, 'revokeObjectURL');
    await expect(prepareWorkbookAssets(source, ['assets/one.png'], signal(), { maxImagePixels: 63 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(create).not.toHaveBeenCalled();
    await expect(prepareWorkbookAssets(source, ['assets/one.png', 'assets/two.png'], signal(), { maxPixels: 127 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(create).toHaveBeenCalledTimes(1); expect(revoke).toHaveBeenCalledWith(create.mock.results[0]!.value);
    const accepted = await prepareWorkbookAssets(source, ['assets/one.png', 'assets/two.png', 'assets/one.png'], signal(), { maxPixels: 128 });
    try { expect(accepted.pixels).toBe(128); } finally { accepted.dispose(); }
  });
  it('rejects changed identities, paths, bytes and over-budget provider responses before making URLs', async () => {
    const source = seed(), create = vi.spyOn(URL, 'createObjectURL');
    for (const edit of [
      (r: WorkbookAssetResponse) => ({ ...r, documentId: 'other' }), (r: WorkbookAssetResponse) => ({ ...r, revision: r.revision + 1 }),
      (r: WorkbookAssetResponse) => ({ ...r, path: 'assets/two.png' }), (r: WorkbookAssetResponse) => ({ ...r, bytes: new Uint8Array(bytes.length) }),
      (r: WorkbookAssetResponse) => ({ ...r, bytes: new Uint8Array(bytes.length + 1) }),
    ]) {
      const provider: WorkbookAssetSource = { ...source, async readAsset(request, signal) { return edit(await source.readAsset(request, signal)); } };
      await expect(prepareWorkbookAssets(provider, ['assets/one.png'], signal(), { maxAssetBytes: bytes.length })).rejects.toBeInstanceOf(Error);
    }
    expect(create).not.toHaveBeenCalled();
  });
  it('isolates mutable provider bytes before asynchronous hashing', async () => {
    const source = seed(); let returned: Uint8Array | undefined;
    const provider: WorkbookAssetSource = { ...source, async readAsset(request, signal) {
      const response = await source.readAsset(request, signal); returned = response.bytes; return response;
    } };
    const digest = crypto.subtle.digest.bind(crypto.subtle);
    vi.spyOn(crypto.subtle, 'digest').mockImplementation((algorithm, data) => { returned!.fill(0); return digest(algorithm, data); });
    const lease = await prepareWorkbookAssets(provider, ['assets/one.png'], signal());
    try { expect(new Uint8Array(await (await fetch(lease.url('assets/one.png')!)).arrayBuffer())).toEqual(bytes); }
    finally { lease.dispose(); }
  });
  it('cancels promptly and revokes earlier URLs even when a provider ignores its signal', async () => {
    const source = seed(), controller = new AbortController(), create = vi.spyOn(URL, 'createObjectURL'), revoke = vi.spyOn(URL, 'revokeObjectURL');
    let late!: (value: WorkbookAssetResponse) => void, started!: () => void;
    const pending = new Promise<void>(resolve => { started = resolve; });
    const provider: WorkbookAssetSource = { ...source, readAsset(request, signal) {
      if (request.path === 'assets/one.png') return source.readAsset(request, signal);
      return new Promise(resolve => { late = resolve; started(); });
    } };
    const work = prepareWorkbookAssets(provider, ['assets/one.png', 'assets/two.png'], controller.signal), outcome = expect(work).rejects.toMatchObject({ code: 'OperationCancelled' });
    await pending; expect(create).toHaveBeenCalledTimes(1); controller.abort(); await outcome;
    expect(revoke).toHaveBeenCalledTimes(1);
    late({ documentId: 'display-assets', revision: 0, path: 'assets/two.png', bytes }); await Promise.resolve();
    expect(create).toHaveBeenCalledTimes(1);
    await expect(prepareWorkbookAssets(source, [], controller.signal)).rejects.toMatchObject({ code: 'OperationCancelled' });
  });
  it('does not create a late URL after cancellation while digesting a resource', async () => {
    const source = seed(), controller = new AbortController(), create = vi.spyOn(URL, 'createObjectURL');
    let release!: (value: ArrayBuffer) => void, start!: () => void;
    const started = new Promise<void>(resolve => { start = resolve; }), digest = crypto.subtle.digest.bind(crypto.subtle);
    vi.spyOn(crypto.subtle, 'digest').mockImplementationOnce(() => new Promise(resolve => { release = resolve; start(); }));
    const work = prepareWorkbookAssets(source, ['assets/one.png'], controller.signal), rejected = expect(work).rejects.toMatchObject({ code: 'OperationCancelled' });
    await started; controller.abort(); await rejected; release(await digest('SHA-256', bytes)); await Promise.resolve();
    expect(create).not.toHaveBeenCalled();
  });
  it('releases earlier URLs when the browser cannot allocate another object URL', async () => {
    const source = seed(), original = URL.createObjectURL, revoke = vi.spyOn(URL, 'revokeObjectURL');
    vi.spyOn(URL, 'createObjectURL').mockImplementationOnce(blob => original(blob)).mockImplementationOnce(() => { throw Error('URL allocation failed'); });
    await expect(prepareWorkbookAssets(source, ['assets/one.png', 'assets/two.png'], signal())).rejects.toThrow('URL allocation failed');
    expect(revoke).toHaveBeenCalledTimes(1);
  });
});
