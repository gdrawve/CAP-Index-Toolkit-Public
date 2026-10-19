# CAP Index Ribbon Design

Ribbon order: Crime Data Prep, Analysis Scope, Classification, Geography, Date & Time, Crime Analysis, Reporting & Export, Excel Essentials, Help & About.

CAP tool buttons use a consistent vector icon family with navy outlines, pale fills and lime accents. Icons are rendered with transparent backgrounds at 16 and 32 pixels. White outer outlines preserve contrast on dark backgrounds; navy shapes remain visible on light backgrounds. The circular logo has a light backing. Large tool buttons use 32-pixel artwork; smaller buttons and dropdown entries use 16-pixel artwork. Existing native Excel commands retain Office's icons.

Help & About contains one large CAP Index dropdown. Its circular logo mark is taken from the official full-color logo on [CAP Index's website](https://capindex.com/), using its left-hand circular mark in a square SVG viewport. The original full logo is preserved under src/ribbon/brand. The menu contains About CAP Index and Help / Tool Guide. About CAP Index opens an embedded branded VBA window with Overview, What's New and Data & Methods tabs. The window displays Grant Drawve (gdrawve@capindex.com) and has website, email compose, Tool Guide and Close buttons. No email is sent automatically. The Tool Guide button requires a writable data workbook that permits new sheets.

Logo source: https://capindex.com/wp-content/uploads/2025/02/CAP-Index-Full-Color-Logo.png

Custom SVG source, transparent PNGs and manifest are under src/ribbon/icons. The build embeds PNGs inside the .xlam and writes image relationships under customUI/_rels/customUI14.xml.rels. No external image files, network calls, image-loading macros or installation steps are required at runtime. Ownership of CAP's logo and new toolkit icon artwork remains with CAP Index.

Create-RibbonIcons.cjs regenerates assets from vector definitions. Its renderer is Sharp in the Codex bundled Node runtime; adjust the renderer path for another build workstation. The generated PNGs are sufficient to build the add-in without running that script. See ribbon-icon-preview.png and ribbon-icon-preview-dark.png for light/dark previews. The About logo is embedded in the UserForm during the build; its source BMP is not needed on a recipient's computer.

The build checks ribbon order, Help menu membership, decoded icon size/content/transparency, embedded parts and relationship references, then reopens the packaged add-in. Native Excel ribbon display remains a pilot acceptance check.
