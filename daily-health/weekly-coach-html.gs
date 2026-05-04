// =============================================================================
// MARKDOWN TO HTML CONVERTER — wraps the AI-generated markdown in the email
// shell, threads inline charts into the right sections (Training, Recovery),
// and supports a two-column layout for "last week vs all time" chart pairs.
// =============================================================================
function markdownToHtml(markdown, weekLabel, inlineImages) {
  // Map section title keywords → chart layout to inject after the header.
  // Matching is case-insensitive substring on the header text (emojis are fine).
  // A mapping has either `cids` (single column, stacked) or `columns` (two-column
  // table layout, each column its own header + ordered cid list).
  const sectionChartMap = [
    { match: 'training', cids: ['fitnessChart'] },
    {
      match: 'recovery',
      columns: [
        { title: 'Last week', cids: ['readinessChart', 'sleepChart', 'vitalsChart'] },
        { title: 'All time',  cids: ['allTimeWeightChart', 'allTimeSleepChart', 'allTimeVitalsChart'] },
      ],
    },
  ];

  function imgTag(cid) {
    return `<img src="cid:${cid}" style="max-width:100%; display:block; margin: 0 auto;">`;
  }

  function chartsForHeader(title) {
    const lc = title.toLowerCase();
    const mapping = sectionChartMap.find(m => lc.includes(m.match));
    if (!mapping) return '';

    if (mapping.columns) {
      // Two-column <table> — most reliable cross-client way to get side-by-side
      // images in HTML email. Each column drops out cleanly if its charts are
      // missing (e.g. fresh history with no all-time data yet).
      const cells = mapping.columns.map(col => {
        const imgs = col.cids
          .filter(cid => inlineImages && inlineImages[cid])
          .map(imgTag)
          .join('');
        if (!imgs) return '';
        const heading = col.title
          ? `<div style="text-align:center; font-weight:600; color:#666; font-size:13px; margin-bottom:4px;">${col.title}</div>`
          : '';
        return `<td style="width:50%; vertical-align:top; padding: 0 2px;">${heading}${imgs}</td>`;
      }).filter(Boolean);
      if (cells.length === 0) return '';
      return `<table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width:100%; border-collapse:collapse; margin: 4px 0 8px 0;"><tr>${cells.join('')}</tr></table>`;
    }

    const imgs = (mapping.cids || [])
      .filter(cid => inlineImages && inlineImages[cid])
      .map(imgTag)
      .join('');
    return imgs ? `<div style="margin: 4px 0 8px 0;">${imgs}</div>` : '';
  }

  const headerReplacer = (_, title) =>
    `</div><h3 style="color:#1a73e8; margin: 24px 0 8px 0;">${title}</h3>${chartsForHeader(title)}<div style="margin-bottom: 16px;">`;

  let html = markdown
    // 1. Format Headers (catches "### Header" or just "**Header**" on its own line)
    .replace(/^(?:###|##)\s*(?:\*\*)?(.*?)(?:\*\*)?$/gm, headerReplacer)
    .replace(/^\*\*(.*?)\*\*$/gm, headerReplacer)

    // 2. Bold text
    .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')

    // 3. Lists (Handles * or - with ANY number of spaces)
    .replace(/^[\*\-]\s+(.+)$/gm, '<li style="margin-bottom: 6px;">$1</li>')
    .replace(/^(\d+)\.\s+(.+)$/gm, '<li style="margin-bottom: 6px;">$2</li>')

    // 4. Italics
    .replace(/\*(.*?)\*/g, '<em>$1</em>');

  // Wrap consecutive <li> elements in a proper <ul> block
  // Extracts only the <li> tags and drops interstitial whitespace (like \n\n)
  // to prevent it from being parsed into broken </div><div> blocks later.
  html = html.replace(/(?:<li[^>]*>.*?<\/li>\s*)+/g, match => {
    const lis = match.match(/<li[^>]*>.*?<\/li>/g).join('');
    return `<ul style="margin: 8px 0 16px 0; padding-left: 20px;">\n${lis}\n</ul>\n\n`;
  });

  // 4. Handle spacing
  html = html.replace(/\n\n/g, '</div><div style="margin-bottom: 16px;">');
  html = html.replace(/\n/g, '<br>');
  html = html.replace(/<\/div><br>/g, '</div>');
  html = html.replace(/<br><div/g, '<div');
  html = html.replace(/<\/ul><br>/g, '</ul>');
  html = html.replace(/<br><ul/g, '<ul');
  // Strip stray <br>s around list items (left over from \n → <br> inside <ul> blocks)
  html = html.replace(/<br><li/g, '<li');
  html = html.replace(/<\/li><br>/g, '</li>');
  // Collapse empty paragraph divs produced when a header sits next to a blank line
  html = html.replace(/<div style="margin-bottom: 16px;"><\/div>/g, '');

  return `<html><head><meta charset="UTF-8"></head><body>
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; max-width: 680px; margin: 0 auto; padding: 20px; color: #333;">
      <div style="background: linear-gradient(135deg, #1a73e8, #34a853); padding: 24px; border-radius: 12px; color: white; margin-bottom: 24px;">
        <h1 style="margin: 0; font-size: 24px;">Weekly Health Report</h1>
        <p style="margin: 8px 0 0; opacity: 0.9; font-size: 16px;">${weekLabel}</p>
      </div>
      <div style="line-height: 1.6; font-size: 15px;">
        <div>${html}</div>
      </div>
      <hr style="border: none; border-top: 1px solid #e0e0e0; margin: 32px 0 16px;">
      <p style="color: #999; font-size: 12px;">Prompted by ${(ATHLETE_PROFILE && ATHLETE_PROFILE.name) || 'you'} · Your numbers, your story.</p>
    </div></body></html>`;
}
