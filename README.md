# kicad-gerberzipper2
KiCad Plot&amp;Zip script plugin (IPC API)

これは GerberZipper の IPC API 対応版です。
まだ開発中であり、充分な機能が実現されていません。

* これまでの SWIG Python API ではなく KiCad で今後サポートされる IPC API および コマンドラインの kicad-cli を使用します。
* Windows KiCad 9.0.3 でテスト。[Preferences]-[Preferences...] 内 [Plunins] タブで [Enable KiCad API] を有効にする必要があります。
* installtest.bat で KiCad 9.0 のプラグインディレクトリに直接インストールします。

旧版との違い
* IPC API から KiCad 本体の言語設定を取得する機能がないため言語が固定。
* ドリルレポートファイル (*.rpt) 作成機能がない。
* 詳細な設定等
  - "ForcePlotInvisible" 設定がない。この機能は KiCad 本体からも 9.0.1 以降で削除される。
  - "ExcludeEdgeLayer" 設定がない。この機能は KiCad 本体で既に廃止。最近の KiCad は同等以上の柔軟性のある"PlotOnAllLayers"がある。
  - "ExcludePadsFromSilk" 設定がない。この機能は KiCad 本体で既に廃止された Fab 用の機能。最近の KiCad は代わりに"Sketch pads on fabrication layers"が設定できる。
  - "DoNotTentVias" 設定がない。この機能は PCB 全体の Via のテントをOn/Offするものだったが既に廃止。現在は Via 個別に制御できる。
  - "LineWidth" 設定がない。この設定は既に廃止されている。経緯が古すぎて良くわからない。多分 Ver 6.x の頃。
  - BOM/POS 未実装。
  


