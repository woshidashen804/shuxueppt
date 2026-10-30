# OOXML 结构速查

调试 pptx 公式问题时的速查表。

---

## 一、公式容器层级

### 内联公式（一段话中间）

```xml
<a:p>
  <a:r><a:rPr sz="2400"/><a:t>当 </a:t></a:r>
  <a14:m>
    <m:oMath>
      <m:r><a:rPr sz="2400" i="1"/><m:t>x</m:t></m:r>
      <m:r><m:rPr><m:sty m:val="p"/></m:rPr><a:rPr sz="2400"/><m:t>→</m:t></m:r>
      <m:r><a:rPr sz="2400"/><m:t>0</m:t></m:r>
    </m:oMath>
  </a14:m>
  <a:r><a:rPr sz="2400"/><a:t> 时</a:t></a:r>
</a:p>
```

`a14:m` 是 DrawingML 2010 的公式载体，挂 `m:oMath`。**这是最稳的形式**，两个引擎都能渲染。

### 独立成行的公式

```xml
<a:p>
  <a14:m>
    <m:oMathPara>
      <m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>
      <m:oMath>…</m:oMath>
    </m:oMathPara>
  </a14:m>
</a:p>
```

`m:oMathPara` 的合法子元素**只有** `m:oMathParaPr` 和 `m:oMath`。里面再放 `a14:m` 就是非法嵌套（缺陷 M-BADNEST）。

### 合法但不推荐的嵌套

```xml
<!-- schema 合法，但两个引擎实测下来 a14:m > oMath 更稳 -->
<a14:m>
  <m:oMathPara>
    <m:oMath>…</m:oMath>
  </m:oMathPara>
</a14:m>
```

---

## 二、mc:AlternateContent 语义

```xml
<mc:AlternateContent xmlns:mc="…">
  <mc:Choice Requires="a14">
    …只有支持 a14 的阅读器才看这里…
  </mc:Choice>
  <mc:Fallback>
    …其他阅读器看这里…
  </mc:Fallback>
</mc:AlternateContent>
```

规则：
- `Requires="a14"` 表示"要 a14 能力才能正确读"。
- 拿不到 a14 的阅读器**必须**用 Fallback。
- **Fallback 为空 = 该阅读器什么都看不到。**

两种放置层级：

| 层级 | 父节点 | Fallback 通常是 |
|---|---|---|
| 形状级 | `p:spTree` / `p:grpSp` | 空白 `p:sp`（`\xa0`）→ **危险** |
| 段落级 | `a:p` | 同一段内容的纯文本镜像 → 安全 |

段落级是 PowerPoint 自己写的正常形式。形状级且 Fallback 空白，就是 WPS 丢内容的根源。

`Requires="p14"` 用于幻灯片切换（`p:transition`），与公式无关，不要动。

---

## 三、m:oMath 内的元素

| 元素 | 含义 | 子元素 |
|---|---|---|
| `m:r` | 数学 run | `a:rPr` + `m:t`（**不是 `a:t`**） |
| `m:t` | 文本 | — |
| `m:f` | 分式 | `m:fPr?` `m:num` `m:den` |
| `m:sSup` | 上标 | `m:sSupPr?` `m:e` `m:sup` |
| `m:sSub` | 下标 | `m:sSubPr?` `m:e` `m:sub` |
| `m:sSubSup` | 上下标 | `m:e` `m:sub` `m:sup` |
| `m:rad` | 根式 | `m:radPr?` `m:deg` `m:e` |
| `m:d` | 定界符 | `m:dPr?` `m:e+` |
| `m:limLow` | 极限 | `m:limLowPr?` `m:e` `m:lim` |
| `m:limUpp` | 上极限 | `m:e` `m:lim` |
| `m:func` | 函数 | `m:funcPr?` `m:fName` `m:e` |
| `m:nary` | 求和/积分 | `m:naryPr?` `m:sub` `m:sup` `m:e` |
| `m:eqArr` | 方程组 | `m:eqArrPr?` `m:e+` |
| `m:m` | 矩阵 | `m:mPr?` `m:mr+` |

**属性块命名**：`m:fPr`、`m:sSupPr`、`m:dPr` 等，以 `Pr` 结尾。属性块内的 `m:ctrlPr` 存字号/字体等格式。

**`m:sty`**：`<m:rPr><m:sty m:val="p"/></m:rPr>` 把该 run 设为正体（`p` = plain），用于函数名 `sin`、`lim`、微分号 `d`。

---

## 四、几何与文本度量

| 单位 | 换算 |
|---|---|
| EMU | `1 英寸 = 914400 EMU` |
| 磅（pt） | `1 英寸 = 72 pt` |

- `a:off/@x,@y`：左上角，EMU
- `a:ext/@cx,@cy`：宽高，EMU
- 组形状 `a:chOff/@x,@y` + `a:chExt/@cx,@cy`：子元素坐标系
  - `off == chOff` 且 `ext == chExt` 时缩放为 1:1，子元素按自身坐标渲染
  - 不一致就会整体缩放，是"排版错乱"的常见原因之一

**文本真实范围**（COM）：`TextRange2.BoundLeft/BoundTop/BoundWidth/BoundHeight`，单位磅。与形状 `Left/Width` 常有很大差异，对齐必须用前者。

**`wrap="none"` + `spAutoFit`**：宽度不可信，PowerPoint 保持中心不变、把框收缩到文字宽度。

---

## 五、动画时间轴

```xml
<p:timing>
  <p:tnLst>
    <p:par><p:cTn id="1" nodeType="tmRoot">
      <p:childTnLst>
        <p:seq><p:cTn nodeType="mainSeq">
          <p:childTnLst>
            <p:par><p:cTn presetClass="entr" nodeType="clickEffect">
              …
              <p:spTgt spid="8"/>   <!-- 指向形状 id -->
            </p:cTn></p:par>
```

- `presetClass`：`entr` 进入 / `exit` 退出 / `emph` 强调 / `path` 路径。
- `nodeType`：`clickEffect` 点击触发 / `withEffect` 同时 / `afterEffect` 之后。
- 形状 `hidden="1"` + `entr` 动画 = 点击后出现。
- `p:spTgt/@spid` 必须能在同页 `p:cNvPr/@id` 里找到，否则是孤儿引用。

---

## 六、最小验证清单

拿到一个可疑 slide，按顺序查：

```
1. XML 能解析吗                    lxml.fromstring
2. 形状 id 有重复吗                p:cNvPr/@id 计数
3. 动画目标都存在吗                spTgt ∈ cNvPr ids
4. 有 Fallback 为空的 AlternateContent 吗
5. oMathPara 里有 a14:m 吗
6. oMath 里有 a:r / a:t 吗（会被 PowerPoint 拒开）
7. 形状越出画布吗                  x+w ≤ 13.33, y+h ≤ 7.5
8. hidden 形状都有动画指向吗
```

八条全过，再用两个引擎实际渲染确认。

---

## 七、脚本调试技巧

**看真实 XML**（不要靠打印猜测）：
```python
print(etree.tostring(sp, pretty_print=True).decode()[:2500])
```

**按形状名定位**：
```python
for sp in root.iter(P + 'sp'):
    cnv = sp.find('.//' + P + 'cNvPr')
    if cnv is not None and cnv.get('name') == 'TextBox 7':
        ...
```

**统计时注意子串误匹配**：
```python
raw.count('m:oMath')                    # 错：也匹配 m:oMathPara
len(re.findall(r'<m:oMath[ >]', raw))   # 对
```

**确认是哪一步改坏的**：把每个变换单独作用在备份副本上，每次只加一步，逐步打印目标指标。不要一次性跑完整流程再看结果。
