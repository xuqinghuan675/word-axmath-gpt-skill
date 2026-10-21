# Security and document handling

This repository automates operations on potentially sensitive Microsoft Word documents through locally installed Word and AxMath software.

- Treat each input document and extracted formula as potentially confidential. Do not attach unredacted customer DOCX, conversion output, logs or license information to public GitHub issues.
- Reports should use synthetic reproductions, or logs redacted to remove document text, local user paths, personal information and product/license keys.
- Keep frozen-source integrity checks, the refusal to overwrite the input, and explicit per-page approval requirements intact when changing the conversion flow.
- Do not disable process ownership checks or attach to an arbitrary existing Word instance to make a conversion appear to work.
- Microsoft Word, AxMath, and third-party Python libraries are not maintained or licensed by this project; keep them updated according to their vendors' guidance.

For a suspected vulnerability, avoid posting exploit details or private documents in a public issue. First use the maintainer contact/reporting options currently available on the repository.

No guarantee of confidentiality, a response-time SLA, or successful conversion is implied by this document.
