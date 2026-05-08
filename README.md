# IC Fitness

Next.js coach workspace with Supabase Auth, protected routes, a live IC Fitness
exercise catalog, and a shared workout template builder.

## Setup

1. Copy `.env.example` to `.env.local`.
2. In Supabase, open Project Settings > API and copy the publishable key into
   `NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY`.
3. Install dependencies and start the app:

```bash
npm install
npm run dev
```

## Auth

- `/auth/login` signs existing users in.
- `/auth/sign-up` creates users and stores `full_name` in the profile row.
- `/dashboard`, `/exercises`, `/templates`, and `/profile` are protected by
  `proxy.ts` and redirect signed-out visitors to `/auth/login`.
- `public.profiles` is linked to `auth.users` and protected with RLS.

## Database

The Auth profile migration is saved at
`supabase/migrations/202605060001_setup_auth_profiles.sql`.

The coach build migrations are saved in `supabase/migrations/`.

The imported catalog lives in the `icfitness` schema. Public security-invoker
views named `icfitness_*` are used by the app so Supabase RLS still controls
access while the frontend can query through the standard Data API.

## Coach Workspace

- `/dashboard` shows catalog/template metrics and recently updated exercises.
- `/exercises` supports live catalog search, filters, detail review, and
  coach/admin create, edit, and hard-delete actions.
- `/exercise-admin` keeps the review queue for missing media, coaching notes,
  equipment, and tracking cleanup.
- `/templates` supports shared template search, create, edit, duplicate, and
  delete actions with ordered workout blocks.
- `/admin` lets admins update user roles.
- Athlete workout assignment and result logging are intentionally out of v1.
