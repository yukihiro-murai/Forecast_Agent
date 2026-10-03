/**
 * Ai.js — A-4 AI 調査の、Vertex AI への問い合わせ。
 *
 * 旧来の A-4（runVertexAIResearch）は、4 つの話題ごとに Vertex AI へ 3〜5 回ずつ問い合わせ、全部で 6 分を超えることがある。
 * 旧来の計算は変えられないので、旧来の UrlFetchApp の代わりにこの道具を渡す:
 *   - 受け取った答えは、その回（actionId）の控えとしてキャッシュに置く
 *   - 時間の区切り（APP_AI_CALC_BUDGET_MS）を過ぎたら、新しい問い合わせはせずに止める（旧来の側は失敗として先へ進むが、その結果は使わない）
 *   - 計算用ブックを組み立て直してもう一度動かすと、同じ問い合わせは控えから答える（時刻と乱数は固定なので、問い合わせの中身も同じになる）
 * 問い合わせ先は Vertex AI（aiplatform）と Vertex AI Search（discoveryengine）だけ。
 */
const APP_AI_CALC_BUDGET_MS = 3 * 60 * 1000;   // 1 回の計算で新しく問い合わせてよい時間（1 回の問い合わせに 1 分ほどかかっても 6 分に収まる）
/** 区切りの時刻（テストで差し替える） */
function appAiDeadline_(startMs) { return startMs + APP_AI_CALC_BUDGET_MS; }
const APP_AI_MAX_ATTEMPTS = 12;               // 組み立て直して動かす回数の上限
const APP_AI_ALLOWED_URL = /^https:\/\/([a-z0-9-]+-)?(aiplatform|discoveryengine)\.googleapis\.com\//;

/** 旧来の UrlFetchApp.fetch の代わり。fetch(url, options) と、止めたかどうか・回数を返す */
function appAiFetcher_(actionId, deadlineMs) {
  const st = { hits: 0, live: 0, stopped: false };
  const response = (code, text) => ({
    getResponseCode: () => code,
    getContentText: () => text,
    getHeaders: () => ({})
  });
  const fetch = (url, options) => {
    url = String(url || '');
    if (!APP_AI_ALLOWED_URL.test(url)) throw new Error('Vertex AI 以外へは問い合わせません: ' + url.split('?')[0].slice(0, 80));
    const body = options && options.payload !== undefined ? String(options.payload) : '';
    const key = 'AIF_' + appSha256Hex_(String(actionId) + '\n' + url + '\n' + body).slice(0, 40);
    const saved = appJobGetResult_(key);
    if (saved.found) { st.hits++; return response(saved.value.code, saved.value.text); }
    if (st.stopped || (st.live > 0 && new Date().getTime() > deadlineMs)) {   // 1 回に 1 つは必ず問い合わせる（必ず先へ進む）
      st.stopped = true;
      throw new Error('時間の区切りのため、続きは次の回に問い合わせます。');
    }
    st.live++;
    const res = UrlFetchApp.fetch(url, options);
    const code = res.getResponseCode();
    const text = res.getContentText() || '';
    // 一時的な失敗（混み合い・Google 側のエラー）は控えない（次の回にもう一度問い合わせる）
    if (code !== 429 && code < 500) appJobPutResult_(key, { code: code, text: text });
    return response(code, text);
  };
  return { fetch: fetch, stopped: () => st.stopped, stats: () => ({ hits: st.hits, live: st.live, stopped: st.stopped }) };
}
