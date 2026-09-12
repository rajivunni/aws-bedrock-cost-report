# Security and data handling

This is a local reporting demo, not an authenticated reporting service. No credentials are included. Use an explicitly selected, least-privilege AWS CLI profile with short-lived credentials stored outside the repository. Review billing visibility and Cost Explorer charges before a live run.

Keep real account lists, generated workbooks, JSON snapshots, credential files and AWS configuration private. Account IDs are identifiers rather than authentication secrets, but can still be confidential. The included fixture and CSV are synthetic. Review staged files before publishing, since `.gitignore` cannot detect every sensitive filename.

The application logs controlled validation messages and exception categories, not raw AWS errors or payloads. Excel service labels are written as literal text to avoid formula interpretation. The report is a static snapshot and does not claim invoice completeness or production security certification.

Do not open a public issue containing credentials or billing data. If a credential is exposed, revoke or rotate it at its issuer. Deleting a file from the latest commit does not remove historical exposure.
