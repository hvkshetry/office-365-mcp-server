const { describe, expect, it } = require('@jest/globals');
const { renderEmailBody } = require('../email/index');

// Outlook HTML bodies routinely run 50-70 KB of style blocks and MSO conditionals around a
// few hundred words. Returning that raw made ordinary "read this email" calls overflow the
// caller's token budget and fail on message size alone.
describe('renderEmailBody', () => {
  it('strips Outlook chrome down to readable text', () => {
    const html = [
      '<html><head><meta http-equiv="Content-Type" content="text/html">',
      '<style type="text/css" style="display:none"><!-- p {margin-top:0;margin-bottom:0} --></style>',
      '</head><body dir="ltr">',
      '<div class="elementToProof" style="font-family:Aptos,Calibri,sans-serif; font-size:12pt">Dear All,</div>',
      '<div><br></div>',
      '<div>Kindly review the attached&nbsp;draft before I send it on.</div>',
      '<p>Best Regards</p>',
      '</body></html>',
    ].join('');

    const out = renderEmailBody({ contentType: 'html', content: html });

    expect(out).toContain('Dear All,');
    expect(out).toContain('Kindly review the attached draft before I send it on.');
    expect(out).toContain('Best Regards');
    // No markup, no style rules, no entities survive.
    expect(out).not.toMatch(/<[a-z/!]/i);
    expect(out).not.toContain('font-family');
    expect(out).not.toContain('&nbsp');
    expect(out.length).toBeLessThan(html.length / 2);
  });

  it('drops script, style and head blocks entirely', () => {
    const html = '<head><style>.x{color:red}</style></head><body><script>alert(1)</script><p>Visible</p></body>';
    const out = renderEmailBody({ contentType: 'html', content: html });
    expect(out).toContain('Visible');
    expect(out).not.toContain('alert');
    expect(out).not.toContain('color:red');
  });

  it('decodes the common entities', () => {
    const out = renderEmailBody({
      contentType: 'html',
      content: '<p>A &amp; B &lt;tag&gt; &quot;quoted&quot; it&#39;s</p>',
    });
    expect(out).toBe('A & B <tag> "quoted" it\'s');
  });

  it('preserves line structure from block and break tags', () => {
    const out = renderEmailBody({
      contentType: 'html',
      content: '<p>Line one</p><p>Line two</p><div>Line three<br>Line four</div>',
    });
    expect(out.split('\n').filter(Boolean)).toEqual([
      'Line one', 'Line two', 'Line three', 'Line four',
    ]);
  });

  it('collapses runaway blank lines rather than emitting a wall of whitespace', () => {
    const out = renderEmailBody({
      contentType: 'html',
      content: '<p>A</p>' + '<div><br></div>'.repeat(30) + '<p>B</p>',
    });
    expect(out).not.toMatch(/\n{3,}/);
  });

  it('passes plain-text bodies through untouched', () => {
    expect(renderEmailBody({ contentType: 'text', content: 'plain <not html> body' }))
      .toBe('plain <not html> body');
  });

  it('accepts Graph HTML content types without depending on their case', () => {
    expect(renderEmailBody({ contentType: 'HTML', content: '<p>Visible</p>' })).toBe('Visible');
  });

  it('handles missing and empty bodies without throwing', () => {
    expect(renderEmailBody(null)).toBe('(empty body)');
    expect(renderEmailBody({})).toBe('(empty body)');
    expect(renderEmailBody({ contentType: 'html', content: '<div>   </div>' }))
      .toBe('(no text content)');
  });
});
