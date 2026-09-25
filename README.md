# Forecast Agent / Forecast vNext

詳細な構成・データ契約・運用手順は [`README_VNEXT_JA.md`](./README_VNEXT_JA.md) を参照してください。

別端末・別AIエージェントで現在の作業を引き継ぐ場合は、最初に
[`AI_HANDOFF.md`](./AI_HANDOFF.md) を参照してください。現在のライブ版、配備先、
アストラゼネカFY2027の状態、テスト・反映手順、変更禁止境界をまとめています。

## 初期版の信頼境界

Forecast vNext の初期版は、Admin Hub とクライアント年度 book を物理的に分離します。Hub は管理者だけに共有し、従業員には担当クライアントの年度 book だけを共有します。Vertex AI の設定、raw prompt、raw AI evidence、他クライアントの情報はクライアント book に保存しません。

ただし、現行の Sheets + bound Apps Script 方式では、クライアント book のスクリプトは実行ユーザー権限で動きます。そのため、従業員の正規入力を許可しながら hidden sheet を暗号学的に改ざん不能にすることはできません。クライアント内の append-only log は「未信頼のステージング」であり、Hub の正本ではありません。

Admin 所有 trigger は Hub へ取り込む前に、次をサーバー側で再検証します。

- book / client / FY と BOOK_REGISTRY の一致
- Forecast Owner、TEAM、入力者、依頼日時、状態遷移の整合
- evidence type の allowlist と、人間入力への AI metadata 混入禁止
- forecast request の hash、時刻、cutoff 再計算
- submitted plan と SUCCESS forecast の対応、金額算術、理由、12か月配分

違反レコードは Hub へ取り込まず、永続的な例外として管理者へ表示します。この検証は誤操作・通常の改ざんリスクを抑えますが、署名付きの信頼境界そのものではありません。

高保証運用へ移行する場合は、`executeAs=USER_DEPLOYING` の Admin 所有 Web App を唯一の書込口にし、OAuth で呼出者を確認したうえで Hub に直接 append してください。その段階ではクライアントからの local fallback を廃止し、通信失敗時は fail-closed とします。

## 初回運用

1. Legacy book の所有者が Forecast vNext の初期設定を開き、ZAC 実績元 Spreadsheet ID を入力します。
2. private な空Spreadsheetへ、中央配備済みのAdmin専用runtimeとClient専用runtimeをそれぞれbindingし、Admin Hubとimmutable Master Templateを生成します。元のLegacy bookやそのbound scriptはコピーしません。
3. 生成した Admin Hub を開き、「自動運用を有効化」を1回実行します。
4. 社員ポータルでFY、ZACクライアント、関与メンバー氏名を指定して年度 book の作成を依頼します。Forecast Ownerは依頼した社内ユーザーへ自動設定されます。

Adminコードを改修した後は、中央projectへclasp反映してからAdmin Hubの「中央配備版へ更新」を実行します。更新対象は検証済み18ファイルだけで、Hubの履歴・設定・正式計画は置換しません。Client UI/MEMOの改修は管理者限定Template Draftから新しいimmutable Template Releaseとして公開され、既存年度bookは原則固定されます。汎用migration APPLYは停止したままです。従業員テスト前で回答・依頼・予測・計画等が完全に0件、source pinsとruntime SHAが一致する空のPilot Clientだけは、read-only事前判定後に同じURLのままcanonical ACTIVE pairへ更新でき、途中停止時は専用journalから復旧します。

初回展開は2～3 Client、明示承認後のcanaryでも最大5 Clientです。6冊目は30冊負荷試験とrelease承認が完了するまでserver-sideで拒否します。Client fileはAdmin管理のprivate root配下にだけ生成し、共有境界を確認できないfolderは使用しません。

クライアント年度 book の通常表示は `1_ホーム` と `2_予測と計画` の2シートです。内部処理、承認履歴、raw AI metadata は従業員向け画面には表示しません。
