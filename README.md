# ai-dev-portfolio

AI駆動開発の実績をまとめたポートフォリオです。
各制作物は「動くもの（URL/ファイル）」「テスト結果」「README」の3点をそろえています。

## 制作者
- 天野昌人（M AI-studio）
- 強み：資料の整理と、AI出力の検証・修正
- 対応時間：平日 1日2時間

## 制作物一覧

| No | 制作物 | 内容 | 成果物 | 状態 |
|---|---|---|---|---|
| 1 | FAQチャットボット | Difyで構築した架空バイクショップのFAQボット | [01-faq-chatbot](./01-faq-chatbot/) | 制作中 |
| 2 | LP公開 | 架空サービスのLP。AI生成直後と検証・修正後のBefore/After | 02-landing-page | 未着手 |
| 3 | AI出力の検証レポート | 重大度分類（致命的/重要/軽微）と再テスト結果 | 03-verification-report | 未着手 |
| 4 | 小さな業務自動化ツール | サンプル入出力・エラーケース付き | 04-automation-tool | 未着手 |

## 進め方
- 毎日：学習20分 / 制作80分 / 動作確認と作業ログ20分
- 毎週：URLまたはファイル・テスト結果・READMEを必ず残す
- 作業ログ：[docs/worklog](./docs/worklog/)
- 検証チェックリスト：[docs/checklist-v1.md](./docs/checklist-v1.md)

## 確認方法
各フォルダのREADMEに「目的・使い方・確認方法」を記載しています。

## 制作実績①：Dify FAQチャットボット

### プロジェクト概要
架空のバイク販売店「モトリンク福岡」を想定し、顧客からの問い合わせに自動回答するAIチャットボットを制作しました。

### 主な機能
- 店舗の営業時間、点検・整備などの問い合わせに回答
- 30件のFAQをナレッジベースとして登録
- Difyのナレッジ検索とClaude Haiku 4.5を組み合わせた回答生成
- ナレッジに存在しない情報について、事実を捏造しないよう回答を制御

### テスト実績
50問のテストを実施し、修正後の最終テストで50問すべてが期待する回答基準を満たしました。

- FAQに記載された質問：22問
- 曖昧な質問：8問
- FAQにない質問：10問
- 誤回答を誘発する質問：10問

### デモ

[チャットボットを試す](https://udify.app/chat/2QqYFYd1MorRjoHf)

[デモ動画を見る（Googleドライブ）](https://drive.google.com/file/d/1oeChRXDISxAaUvJeab7-nIJfTep5MxE8/view?usp=sharing)

### 使用技術
Dify / Claude Haiku 4.5 / Gemini Embedding / RAG（検索拡張生成）

※架空店舗を題材とした自主制作プロジェクトです。
