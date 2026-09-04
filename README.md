# Willow Park Checks V2.5 — Private PDF Cloud Archive

This version stores assessments locally in IndexedDB and, when the tablet is connected to the Willow Park Supabase account, automatically generates and uploads a readable PDF to the private `risk-assessments` bucket on submission.

Cloud path format:
`<Academic Year>/<Term>/<Area>/<DD-MM-YYYY - Area - Staff Member.pdf>`

The Supabase project URL and publishable key are frontend-safe values. Never put a Supabase secret/service-role key or the cloud account password in this repository.

## Tablet setup
1. Open Settings → PDF cloud archive.
2. Enter the Willow Park cloud account email/password.
3. Tap Connect this tablet once.
4. Complete a test assessment and verify the PDF appears in Supabase Storage → risk-assessments.

If offline or cloud upload fails, the assessment remains stored locally and can be retried from the completion screen or Settings → Upload pending assessments.

JSON backup remains available for disaster recovery only.
