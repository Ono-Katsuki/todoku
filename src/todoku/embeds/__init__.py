"""文書の中に埋め込まれたものを、平文で読む（読むだけ。書き込みはしない）。

- chart：グラフ（Excel・Word の DrawingML のグラフ c:chartSpace、新しい種類のグラフ cx:chartSpace）を、
  題名つきのデータの表にする。値はグラフが持っている値の控え（c:strCache・c:numCache）から読み、控えがなければ参照先のセルを読む
- biffchart：古い Excel（.xls）のグラフ（BIFF8 のグラフの部分 BOF 0x0020：シートに置いたグラフ・グラフのシート）
- smartart：SmartArt（dgm:dataModel）の字を、つながりの順の入れ子の一覧にする
- ole：OLE の部品（Word・Excel・PowerPoint・PDF・数式・パッケージ・その他）を見分けて、中身を平文にする
- ppt：PowerPoint のスライドの字（.pptx の部品と、古い形の PowerPoint Document の字の atom）
- mtef：数式エディター 3.0・MathType の数式（Equation Native の MTEF）の字
- excel・word：Excel のシート・Word の文書ごとに、グラフ・SmartArt・埋め込みを集めて窓の形にする
- audit：窓とは別の読み方で、埋め込みの字を全部拾い、窓に出ているかを確かめる（読み落とし）
"""
