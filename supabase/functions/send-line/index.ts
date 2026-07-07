/**
 * send-line Edge Function
 *
 * Broadcasts a message to ALL doctors who have a line_user_id registered.
 * Secrets required (set in Supabase Dashboard → Settings → Edge Functions → Secrets):
 *   LINE_CHANNEL_ACCESS_TOKEN  — your Messaging API channel access token
 *
 * Deploy:
 *   supabase functions deploy send-line --no-verify-jwt
 */

import { serve } from 'https://deno.land/std@0.177.0/http/server.ts';
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2';

const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, content-type',
};

serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS });

  try {
    const { message } = await req.json();
    if (!message) return new Response('missing message', { status: 400, headers: CORS });

    const token = Deno.env.get('LINE_CHANNEL_ACCESS_TOKEN');
    if (!token) return new Response('LINE token not configured', { status: 500, headers: CORS });

    // Fetch all doctors with a LINE user ID from DB
    const supabase = createClient(
      Deno.env.get('SUPABASE_URL')!,
      Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!,
    );
    const { data: doctors } = await supabase
      .from('doctors')
      .select('line_user_id')
      .not('line_user_id', 'is', null);

    const results = await Promise.allSettled(
      (doctors ?? []).map(d =>
        fetch('https://api.line.me/v2/bot/message/push', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            Authorization: `Bearer ${token}`,
          },
          body: JSON.stringify({
            to: d.line_user_id,
            messages: [{ type: 'text', text: message }],
          }),
        }),
      ),
    );

    const sent = results.filter(r => r.status === 'fulfilled').length;
    return new Response(JSON.stringify({ sent }), { headers: { ...CORS, 'Content-Type': 'application/json' } });
  } catch (e) {
    return new Response(String(e), { status: 500, headers: CORS });
  }
});
