# 修复配方

每类修复的 XML 前后对照与实现要点。所有脚本默认 dry-run，`--write` 才落盘。

---

## 通用约定

**命名空间**
```
a    http://schemas.openxmlformats.org/drawingml/2006/main
p    http://schemas.openxmlformats.org/presentationml/2006/main
m    http://schemas.openxmlformats.org/officeDocument/2006/math
a14  http://schemas.microsoft.com/office/drawing/2010/main
mc   http://schemas.openxmlformats.org/markup-compatibility/2006
```

**写回方式**：不要用 zipfile 直接覆盖成员（会丢压缩属性）。标准做法是读全部条目写到临时 zip，再整体替换。

```python
zin = zipfile.ZipFile(src)
tmp = tempfile.mkdtemp()
out = os.path.join(tmp, 'o.zip')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith('ppt/slides/slide') and item.filename.endswith('.xml'):
            root = etree.fromstring(data)
            if transform(root):
                data = etree.tostring(root, xml_declaration=True,
                                      encoding='UTF-8', standalone=True)
        zout.writestr(item, data)
zin.close()
shutil.copyfile(out, src)
```

**序列化参数**必须是 `xml_declaration=True, encoding='UTF-8', standalone=True`，否则 PowerPoint 可能拒绝打开。

**不变量断言**：每个脚本在写回前比较该页 `<m:oMath[ >]` 计数，变化即 `SystemExit`。这是防止修一个坏两个的唯一可靠手段。

---

## 1. 拆非法 oMathPara 嵌套（真正修复 WPS 缺内容）

**变换**
```xml
<!-- 前 -->
<a14:m><m:oMathPara><a14:m><m:oMath>…</m:oMath></a14:m></m:oMathPara></a14:m>

<!-- 后：外层 a14:m 直接持有 oMath -->
<a14:m><m:oMath>…</m:oMath></a14:m>
```

**实现**：遍历所有 `a14:m`，若其直接子元素含 `m:oMathPara`，且该 `oMathPara` 下有 `a14:m/m:oMath`，则清空外层子节点、只挂 `oMath`。

**不要**把 `oMathPara` 提到 `a14:m` 外面（`<a14:m><m:oMathPara><m:oMath>`），虽然 schema 上合法，但逐处实测下来 `a14:m > oMath` 在两个引擎里都最稳。

**这是唯一被实测证明能让 WPS 恢复公式的修复**（5 份课件对照，见缺陷目录的表格）。

---

## 1·补、展开空 Fallback（对 WPS 无效，仅作兼容保险）

**判定**：父节点是 `spTree` 或 `grpSp`，`Choice.Requires == "a14"`，`Choice` 的子元素全是 `p:sp`，`Fallback` 内无文字。

**变换**
```xml
<!-- 前 -->
<mc:AlternateContent>
  <mc:Choice Requires="a14"><p:sp>…内容…</p:sp></mc:Choice>
  <mc:Fallback><p:sp>\xa0</p:sp></mc:Fallback>
</mc:AlternateContent>

<!-- 后：直接把 p:sp 放回原位置 -->
<p:sp>…内容…</p:sp>
```

**实测结果：对 PowerPoint 和 WPS 的逐页墨迹变化均为 0。** WPS 读的是 Choice 分支。这个变换只在目标读者可能是真正走 Fallback 的软件时才有意义，不要用它解释 WPS 症状。

**根元素**：展开后 `a14` 前缀会散落在各个元素上。规范化成 PowerPoint 的写法——

```xml
<p:sld xmlns:a="…" xmlns:r="…" xmlns:p="…"
       xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
       xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
       mc:Ignorable="a14">
```

重建根元素时注意保留原有 `attrib` 与全部子节点，并合并已有的 `mc:Ignorable` 词表（可能已含 `p14`）。

**要点**
- 插入时保持 `Choice` 内子元素的原有顺序，动画 z 序依赖它。
- 只处理 `Fallback` 为空的；`Fallback` 有真实文字的说明该结构是合法降级方案。

---

## 3. lim 重建

**变换**
```xml
<!-- 前 -->
<m:sSub>
  <m:sSubPr><m:ctrlPr>…</m:ctrlPr></m:sSubPr>
  <m:e><m:r><m:t>lim</m:t></m:r></m:e>
  <m:sub><m:r><m:t>x→0</m:t></m:r></m:sub>
</m:sSub>

<!-- 后 -->
<m:limLow>
  <m:limLowPr><m:ctrlPr>…</m:ctrlPr></m:limLowPr>
  <m:e><m:r><m:t>lim</m:t></m:r></m:e>
  <m:lim><m:r><m:t>x→0</m:t></m:r></m:lim>
</m:limLow>
```

**要点**
- 只匹配 `m:e` 文本恰好等于 `lim`（去掉空白后）的情况。
- `ctrlPr` 要搬过去，否则字号/字体丢。
- 替换时用 `list(parent).index(node)` 记原位置再 `insert`，保证文档顺序不变。

---

## 4. 去掉空函数壳

**变换**
```xml
<!-- 前：渲染成一个空参数框 -->
<m:func>
  <m:fName><m:r><m:t>sin</m:t></m:r></m:fName>
  <m:e/>
</m:func>

<!-- 后：函数名作为普通公式 run 保留 -->
<m:r><m:t>sin</m:t></m:r>
```

**只处理** `m:e` 完全为空的情况（无文本、无非属性子元素）。有参数的 `m:func` 不要碰。

---

## 5. 指数合成为真上标

**变换**
```xml
<!-- 前 -->
<a:r><a:rPr sz="2000"/><a:t>d(e</a:t></a:r>
<a:r><a:rPr sz="1400" baseline="30000"/><a:t>−x</a:t></a:r>

<!-- 后 -->
<a:r><a:rPr sz="2000"/><a:t>d(</a:t></a:r>
<a14:m>
  <m:oMath>
    <m:sSup>
      <m:e><m:r><a:rPr i="1" sz="2000"/><m:t>e</m:t></m:r></m:e>
      <m:sup><m:r><a:rPr sz="1400"/><m:t>-x</m:t></m:r></m:sup>
    </m:sSup>
  </m:oMath>
</a14:m>
```

**分数指数**：`3/2` 在 `m:sup` 内做成真正的 `m:f`
```xml
<m:sup>
  <m:f>
    <m:fPr><m:type m:val="bar"/></m:fPr>
    <m:num><m:r><m:t>3</m:t></m:r></m:num>
    <m:den><m:r><m:t>2</m:t></m:r></m:den>
  </m:f>
</m:sup>
```

**致命细节**：`oMath` 内部的文本 run 必须是 `m:r` + `m:t`，**不能是 `a:r` + `a:t`**。写成 `a:r` 会让 PowerPoint 直接拒绝打开文件（报"PowerPoint could not open the file"），而 XML 本身是能解析的，很难一眼看出。

```python
# 正确
run = etree.Element(M + 'r')
rpr = etree.SubElement(run, A + 'rPr')   # 属性块仍是 a:rPr
t = etree.SubElement(run, M + 't')
```

写完必须校验：
```python
for om in root.iter(M + 'oMath'):
    assert not any(True for _ in om.iter(A + 'r'))
    assert not any(True for _ in om.iter(A + 't'))
```

**边界处理**：从前一个 run 里切出携带指数的字符（`e` / `x` / `)`），剩下的文本留在原 run。若前一个 run 因此变空，仍保留它（不要删，删除会改变 run 序列，后续匹配会错位）。

---

## 6. 答案位置修正

坐标换算：`EMU = 914400 × 英寸`。

**关键坑：`wrap="none"` + `spAutoFit` 的形状按中心缩放。**
这种文本框的宽度不可信——PowerPoint 会保持中心不动、把宽度收缩到文字宽度。所以只改 `a:off/@x` 无法把文字挪到位。两种解法：

1. 同时设置 `a:ext/@cx` 使宽度贴合文字：
   ```
   box_x = desired_text_left - (box_w - text_w) / 2
   ```
2. 或把段落对齐改成左对齐（`a:pPr/@algn="l"`）。注意若该段落只有 `a14:m` 公式、没有 `a:r`，`algn` 不一定生效，此时用解法 1。

**对齐依据**：用 PowerPoint 实测的 `TextRange2.BoundLeft/BoundWidth`，不要用形状 `Left/Width`。形状框和文字实际范围经常差很多（例如形状宽 4.81"，文字只有 1.06"）。

**移动前先确认不是点击揭示设计**：
- 查 `p:timing` 里两个形状是不是都是 `presetClass="entr"`；
- 位置完全一致、答案文本包含题干前缀 → 是揭示设计，别动。

---

## 7. 删除冗余画布外形状

前提：题干已包含同样信息（如 `（√）`），画布外形状只是残留。

**顺序很重要**：
1. 从 `p:timing` 里删掉所有指向该 `spid` 的 `p:par` / `p:cond` 节点；
2. 再删 `p:sp`。

反了会留下孤儿动画目标。校验：
```python
ids = {c.get('id') for c in root.iter(P+'cNvPr')}
targets = {t.get('spid') for t in root.iter(P+'spTgt')}
assert not (targets - ids)
```

---

## 8. 变更记录与回滚

每次 `--write` 之前先整目录备份，命名带时间戳：

```powershell
$ts = Get-Date -Format "yyyyMMdd_HHmm"
Copy-Item -Path "<decks>\*.pptx" -Destination "<review>\backup_before_fix_$ts" -Force
```

`Copy-Item` 配 `-Path` 才能用通配符，`-LiteralPath` 不支持。

回滚就是把备份目录下的 pptx 复制回去。

---

## 9. 一次完整修复的推荐顺序

```
备份
  ↓
inspect_pptx.py          出清单，交用户确认
  ↓
fix_invisible.py --nest  拆 oMathPara 非法嵌套（真正修复 WPS 缺内容）
  ↓  verify_integrity.py
fix_limit.py             极限
  ↓
fix_empty_math.py        空函数壳
  ↓
fix_exponent.py          指数
  ↓
audit_overlap.py         版式（人工确认后再动坐标）
  ↓
compare_render.py        双引擎回归验收
  ↓
verify_integrity.py      最终结构校验
```

先修"看不见"，再修"不好看"。前者是内容丢失，后者是风格问题；混在一起改，出问题时分不清是哪个改动导致的。

`fix_invisible.py --unwrap` 不属于主流程。它改动面大、对 WPS 无效果，只在确认目标读者包含老版本 Office 时才追加，且要单独跑一次 `compare_render.py` 确认零回归。

---

## 10. 下结论前必须做的单变量实验

本项目最重要的方法论。任何"改了 X 就能修好 Y"的判断，都要这样验证：

1. 从**未修复的备份**出发（不是当前文件——当前文件里问题可能已经被别的步骤顺带修掉了）。
2. 构造只含单一变换的变体，一次只加一个。
3. 每个变体分别用 PowerPoint 和 WPS 渲染全本，逐页量墨迹。
4. 只有"目标引擎墨迹上升、另一引擎墨迹不变"的那个变体，才是真正的修复。

现成的工具：

```bash
# 构造变体（u=只展开 Fallback，f=只拆嵌套，uf=两者）
python build_variant.py <备份.pptx> <输出.pptx> uf

# 单进程渲染取每页墨迹（避免两个 COM 实例冲突）
python render_ink.py KWPP.Application <文件.pptx> <out.json>
python render_ink.py PowerPoint.Application <文件.pptx> <out.json>
```

**踩过的坑**：一开始判定"空 Fallback 导致 WPS 缺内容"，是因为看到 1265 处可疑结构和 WPS 确实缺内容，就默认了因果。实际跑单变量实验后发现展开 Fallback 的墨迹变化恒为 `+0`，真正的原因是同批文件里 207 处 `oMathPara` 非法嵌套。**结构可疑不等于它导致问题。**
