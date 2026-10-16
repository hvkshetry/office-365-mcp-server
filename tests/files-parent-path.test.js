/**
 * A folder path must address the drive it was given, not the caller's own OneDrive.
 *
 * Two defects sat in the path branch of create_folder and upload. The parameter is documented as
 * "Parent folder path", but its value was interpolated into the request as though it were a Graph
 * endpoint, so passing a path produced a malformed URL and a 400 naming a segment the caller never
 * wrote. And driveId was ignored on that branch entirely, so a caller who correctly supplied a
 * SharePoint drive plus a path had the item created in their personal OneDrive instead -- with the
 * call reporting success.
 */

const { driveParentPathEndpoint } = require('../files');

const DRIVE = 'drive-fixture-1';

describe('driveParentPathEndpoint', () => {
  test('a bare root path addresses the drive root, not /me/drive', () => {
    expect(driveParentPathEndpoint('/', DRIVE)).toBe(`/drives/${DRIVE}/root`);
    expect(driveParentPathEndpoint('', DRIVE)).toBe(`/drives/${DRIVE}/root`);
  });

  test('a folder path is addressed inside the given drive', () => {
    expect(driveParentPathEndpoint('02 Project Documents', DRIVE))
      .toBe(`/drives/${DRIVE}/root:/02%20Project%20Documents:`);
  });

  test('a nested folder path keeps its separators', () => {
    expect(driveParentPathEndpoint('/02 Project Documents/01 Process/', DRIVE))
      .toBe(`/drives/${DRIVE}/root:/02%20Project%20Documents/01%20Process:`);
  });

  test('encodes reserved characters within each folder name', () => {
    expect(driveParentPathEndpoint('Reports #1/50% complete', DRIVE))
      .toBe(`/drives/${DRIVE}/root:/Reports%20%231/50%25%20complete:`);
  });

  test('without a driveId it falls back to the user own drive', () => {
    expect(driveParentPathEndpoint('/', undefined)).toBe('/me/drive/root');
    expect(driveParentPathEndpoint('Reports', undefined))
      .toBe('/me/drive/root:/Reports:');
  });

  test('a value that is already a Graph endpoint is passed through unchanged', () => {
    // Existing callers pass the default '/me/drive/root'; they must keep working.
    expect(driveParentPathEndpoint('/me/drive/root', undefined)).toBe('/me/drive/root');
    expect(driveParentPathEndpoint(`/drives/${DRIVE}/root`, DRIVE))
      .toBe(`/drives/${DRIVE}/root`);
    expect(driveParentPathEndpoint('/sites/abc/drive/root', DRIVE))
      .toBe('/sites/abc/drive/root');
  });

  test('the old behaviour produced a malformed children endpoint', () => {
    // Regression guard, stated as the shape that used to fail: '/' + '/children'.
    const endpoint = `${driveParentPathEndpoint('/', DRIVE)}/children`;
    expect(endpoint).not.toContain('//children');
    expect(endpoint).toBe(`/drives/${DRIVE}/root/children`);
  });
});
