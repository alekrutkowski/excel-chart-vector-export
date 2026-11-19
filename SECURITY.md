# Security policy

This project is maintained on a best-effort basis. For non-public security concerns, contact **alek.rutkowski@gmail.com** rather than posting a proof of concept in a public issue.

## Supported versions

Security fixes are intended for the latest published version; earlier releases are not guaranteed ongoing maintenance.

## Safety considerations

- Avoid processing workbooks from untrusted sources unless Microsoft Excel and all rendering dependencies are fully patched.
- The native backend disables macros and suppresses external-link refresh, but rendering third-party Office documents still has risks.
- Native EMF extraction uses the Windows clipboard and replaces its current contents.
- The included source does not upload workbook contents to a remote service.
- SVG output is inspected for external links, scripting elements and event-handler attributes; treat SVG from untrusted sources as potentially active content until independently checked.
- Never commit a paid Aspose license or sensitive workbooks to a public repository.

Report the project version, affected input format, a minimal synthetic reproduction and a description of the impact. Do not send private spreadsheet data unless expressly agreed.
