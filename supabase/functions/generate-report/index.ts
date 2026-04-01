// supabase/functions/generate-report/index.ts
//
// Deno Edge Function — AI Analyst proxy (multi-turn)
// Verifies Supabase auth, forwards a full message history + data to Claude,
// returns { type, narrative, chart }
//
// Deploy:  supabase functions deploy generate-report
// Secret:  supabase secrets set ANTHROPIC_API_KEY=sk-ant-...
// ─────────────────────────────────────────────────────────────────────────────

import { serve } from 'https://deno.land/std@0.168.0/http/server.ts';

// ── CORS ─────────────────────────────────────────────────────────────────────
const CORS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
  'Access-Control-Allow-Methods': 'POST, OPTIONS',
};

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...CORS, 'Content-Type': 'application/json' },
  });
}

// ── System prompt (permanent — never changes between requests) ────────────────
const SYSTEM_PROMPT = `
You are an AI analyst for Ghostwrite, a collectible sports figures brand.
Your job: analyze market data and answer questions from the Ghostwrite team.
You are having a conversation — prior turns give you full context. Use them.

═══ PRODUCT CONTEXT ═══
• Ghostwrite sells collectible sports figures organised into "sets" (product lines).
• Sets are organised by sport/year, e.g. "NBA 2024-2025 Game Face", "MLB 2025 Game Face".
• Each set contains multiple players, each available in multiple variants (rarity tiers).
• Variant rarity, lowest → highest: Base → Victory → Emerald → Gold → Chrome → Fire
  (some sets also have SP and Bet-On-Women / VIDES variants)
• "Hits" = any non-Base variant (the desirable, valuable figures).
• Figures are sold in sealed "cases" containing a fixed number of figures (figures_per_case).
• Buyers open cases hoping to pull rare hit variants; the economics resemble trading-card boxes.

═══ KEY FORMULAS ═══

1. HIT RATIO
   sets.hit_ratio ≈ 1.0 means roughly 1 hit per case.
   Expected hits per case    = sets.hit_ratio
   Expected base figures     = sets.figures_per_case − sets.hit_ratio

2. POPULATION-WEIGHTED AVG HIT PRICE  (PWAHP)
   PWAHP = Σ(card_price × card_population) / Σ(card_population)
   Only include non-Base variants.
   Use price (last_sale_price if available, else avg_price) from cards array.
   card_population comes from the setMath array.

3. EXPECTED CASE VALUE  (ECV)
   ECV = (hit_ratio × PWAHP) + ((figures_per_case − hit_ratio) × avg_base_price)
   avg_base_price = mean of Base-variant prices for the set.

4. CASE ROI
   ROI = (ECV − case_price) / case_price × 100 %
   Positive → statistically profitable to buy and flip; negative → expected loss.

5. BREAK-EVEN HIT RATIO
   The hit_ratio at which ECV = case_price:
   break_even = (case_price − base_per_case × avg_base_price) / PWAHP

6. RARITY PREMIUM  (same player, two variants)
   premium = hit_price / base_price

═══ DATA STRUCTURE (provided in the SESSION DATA section at the end of this prompt) ═══
{
  sets:     [ { name, league, casePrice, figuresPerCase, hitRatio, retailPrice } ]
  cards:    [ { set, player, variant, price, sales, avg30d, avg90d, sales30d, sales90d } ]
             price = avg_90d (when 90-day sales exist), else avg_price; null if none
             estimated:true flags model gap-fill cards that have no real sales
  setMath:  [ { setName, player, variant, population } ]
  setStats: [ {
    set,
    casePrice,           ← price paid for a case (website "Case Retail" label)
    hitRatio,            ← guaranteed hits per case
    figuresPerCase,      ← total figures in a case
    pwahp,               ← population-weighted avg hit price ($)
    avgBasePrice,        ← mean Base variant price ($)
    ecv,                 ← expected case value at current prices ($)
    roi,                 ← case ROI % — e.g. 41.0 means +41%
    breakEvenHitRatio,   ← hitRatio where ECV = casePrice
    hitCardsWithPrice,   ← # non-Base combos with a price
    baseCardsWithPrice   ← # Base combos with a price
  } ]  ← AUTHORITATIVE — identical to website display, computed with same formulas
  engineParams: { version, tag } | null
}

⚠️  MANDATORY — setStats IS THE ONLY SOURCE FOR CASE ECONOMICS:
• ROI, ECV, PWAHP, break-even, case price, hit ratio → use setStats fields ONLY.
• NEVER aggregate cards[] or setMath[] to compute these — raw arrays are for
  player-level questions only (e.g. price spreads, top players by variant).
• "What is the ROI for set X?" → answer is setStats[X].roi. Do not recalculate.
• Hypotheticals (e.g. "what if hitRatio dropped to 0.5?"): modify only that one
  variable; use setStats.pwahp and setStats.avgBasePrice as fixed price inputs.

═══ OUTPUT FORMAT  (respond with ONLY valid JSON — no markdown fences, no prose outside the object) ═══
{
  "type": "answer" | "question",
  "narrative": "markdown text (use ## headers, **bold**, bullet lists with -)",
  "chart": {
    "type": "bar" | "line" | "doughnut",
    "title": "short chart title",
    "labels": ["Label 1", "Label 2"],
    "datasets": [
      {
        "label": "series name",
        "data": [1.5, 2.3, -0.8],
        "backgroundColor": ["#3ddc84", "#ff4455", "#e8ff47"]
      }
    ]
  }
}
Set "chart" to null when a chart would not add value.

"type" rules:
  "question" — use ONLY when the request is genuinely ambiguous and you cannot produce a
               meaningful answer without one specific piece of information. Ask for the
               single most important missing detail in one sentence ending with "?".
               Always set "chart" to null for questions.
  "answer"   — use for all completed analyses, calculations, and explanations.
               When in doubt, make a reasonable assumption, state it, and answer directly.
               Do NOT ask for clarification if you can infer reasonable intent from context.

═══ CHART GUIDELINES ═══
• ONLY include a chart when the user explicitly requests one — e.g. "show me a chart",
  "make a bar graph", "plot this", "visualize", "graph it", etc.
• For ALL other responses, set "chart" to null — even when a chart would be informative.
• When a chart IS included, keep "narrative" to 1–2 short sentences describing what
  the chart is showing (it appears as a caption directly below the chart).
• Colors: #3ddc84 (green) = positive/profitable, #ff4455 (red) = negative/loss,
          #e8ff47 (accent) = neutral highlight, #555 = secondary data.
• For ROI or % bars: colour each bar individually (green if ≥ 0, red if < 0).
• Keep labels ≤ 20 characters; abbreviate set names if needed.

═══ STYLE ═══
• Be concise and numerical. Every claim should cite a number.
• Use "$" for prices, "%" for rates, "~" for approximations.
• 2–4 short paragraphs or bullet sections maximum.
• If data for a calculation is missing or too sparse, say so explicitly.
• In follow-up turns you already have context — do not re-explain earlier answers.
`.trim();

// ── Model ─────────────────────────────────────────────────────────────────────
const CLAUDE_MODEL = 'claude-sonnet-4-5';

// ── Message type ──────────────────────────────────────────────────────────────
interface Message {
  role: 'user' | 'assistant';
  content: string;
}

// ── Main handler ─────────────────────────────────────────────────────────────
serve(async (req: Request) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: CORS });
  }

  // ── Auth ──────────────────────────────────────────────────────────────────
  // Supabase's gateway verifies the JWT before the request reaches this function.
  // We just confirm a Bearer token is present as a basic sanity check.
  const authHeader = req.headers.get('Authorization');
  if (!authHeader?.startsWith('Bearer ')) return json({ error: 'Unauthorized' }, 401);

  // ── Parse body ────────────────────────────────────────────────────────────
  let messages: Message[];
  let data: unknown;
  try {
    const body = await req.json();
    messages = body.messages;
    data     = body.data;
    if (!Array.isArray(messages) || messages.length === 0) {
      throw new Error('messages array is required');
    }
    if (messages[0]?.role !== 'user') {
      throw new Error('first message must be from user');
    }
  } catch (e) {
    return json({ error: String(e) }, 400);
  }

  // ── Build system prompt (instructions + session data) ────────────────────
  // Data is embedded here — not injected into message[0] — so Anthropic's
  // prompt caching can cache this entire block across turns.
  // Cache TTL: 5 min (refreshed on each hit), so follow-up questions within
  // a session pay ~10% of the data token cost vs 100% without caching.
  const systemText = data
    ? `${SYSTEM_PROMPT}\n\n═══ SESSION DATA ═══\n${JSON.stringify(data)}`
    : SYSTEM_PROMPT;

  // ── Anthropic API key ─────────────────────────────────────────────────────
  const apiKey = Deno.env.get('ANTHROPIC_API_KEY');
  if (!apiKey) return json({ error: 'ANTHROPIC_API_KEY secret is not configured' }, 500);

  // ── Call Claude ───────────────────────────────────────────────────────────
  let claudeRes: Response;
  try {
    claudeRes = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'x-api-key':         apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-beta':    'prompt-caching-2024-07-31',
        'content-type':      'application/json',
      },
      body: JSON.stringify({
        model:      CLAUDE_MODEL,
        max_tokens: 2048,
        system: [
          {
            type: 'text',
            text: systemText,
            cache_control: { type: 'ephemeral' },
          },
        ],
        messages,   // plain conversation — no data injection needed
      }),
    });
  } catch (e) {
    return json({ error: `Failed to reach Anthropic API: ${String(e)}` }, 502);
  }

  if (!claudeRes.ok) {
    const errText = await claudeRes.text();
    return json({ error: `Anthropic API error ${claudeRes.status}: ${errText}` }, 502);
  }

  const claudeBody = await claudeRes.json();
  const rawText: string = claudeBody.content?.[0]?.text ?? '';

  // ── Parse Claude's JSON response ──────────────────────────────────────────
  const cleaned = rawText
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```\s*$/, '')
    .trim();

  let parsed: { type: string; narrative: string; chart: unknown };
  try {
    parsed = JSON.parse(cleaned);
    // Ensure type field always exists
    if (!parsed.type) parsed.type = 'answer';
  } catch (_) {
    parsed = { type: 'answer', narrative: rawText, chart: null };
  }

  return json(parsed);
});
