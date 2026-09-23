# NailVesta 链接数据深度分析

TikTok Shop 运营数据（product analysis list 每日导出）+ 达人组数据（Creators / Products / Videos / LIVE）的深度分析工具。

## 部署到 Streamlit Community Cloud

1. 这个仓库直接 Deploy，Main file 填 `nailvesta_link_analytics_app.py`。**`.streamlit/config.toml` 这个文件也要一起上传**（配色主题在里面，是 GitHub 默认不显示但会上传的隐藏文件，直接拖进网页上传框，或者用 git 推的话本来就会带上）。
2. **部署后必须做**：App 右下角 ⋮ → **Settings → Secrets**，粘贴下面这段，把 `xxx` 换成真实值（**只有你自己能看到这个 Secrets 页面，不会进代码仓库、不会给访问者看到**）：

   ```toml
   LARK_HOST = "Lark 国际版 (larksuite.com)"
   LARK_BASE_URL = "https://qusjdzq12vah.sg.larksuite.com/base/XwZwbY0NTairFWsejC3lr0zjg0t"
   LARK_APP_ID = "xxx"
   LARK_APP_SECRET = "xxx"
   ```

   达人组数据跟运营数据是同一个租户，同一份 App ID/Secret 会自动沿用，不用重复填 `CREATOR_APP_ID`/`CREATOR_APP_SECRET`。

3. 打开 App，侧边栏「① 测试连接」确认能连上，再点「② 🔄 从 Base 拉取全部」和「🔄 拉取达人组数据」。

## 已知限制（公开部署 vs 本地跑的区别）

- **每次容器重启缓存会清空**：Streamlit Cloud 的磁盘不是永久的，应用长时间没人访问会休眠，醒来后本地缓存的文件都没了，得重新点一次拉取（运营数据 185 个文件 + 达人组数据 186 个文件，全新拉取大概各要 3–5 分钟）。本地跑不会有这个问题，缓存永久存在硬盘上。
- **App Secret 安全**：侧边栏的 App ID / App Secret 输入框现在**故意留空**，哪怕 Secrets 里配置好了也不会回显——因为 Streamlit 的密码框只是界面上显示圆点，真实值还是会发到浏览器，公开给别人访问的话谁都能在控制台里看到。现在这样改了之后，配置好 Secrets 就会在服务器这边悄悄生效，输入框全程看不到任何真实值。

## 页面结构（2026-09-22 改版）

不再是横排 tab 一长条页面，改成侧边栏多页导航（`st.navigation`），点哪个 section 就单独打开那一页、自动回到顶部，不用再上下滑着找：
- **总览**：核心指标卡片 + 最近7天一句话结论 + 快捷入口 + 大盘趋势图
- **深度分析**：两段对比·归因／多粒度对比／前端指标／后端指标
- **监控与溯源**：异常检测与归因／链接下钻／日环比链条&回暖／改动效果溯源
- **参考**：指标归档
- **达人组**（Creators/Products/Videos/LIVE，拉了这部分数据才会出现，就算没拉运营数据也能单独看）：
  - **总览**：四张表整体情况 + GMV 趋势 + LAST CHANCE 占比等三个重点追问
  - **全字段目录**：达人表格里能拿到的每一个字段（不只是必须要看的那几个），自动归到前端(曝光/互动/内容产出)或后端(成交/成本/售后)，按表按分类画趋势
  - **排行榜**：Creators/Products/Videos/LIVE 四张表各自的 Top 20 排行，GMV/播放量/互动率/完播率/佣金成本任选
  - **内容与互动质量**：内容生产节奏、完播率/互动率/观看时长趋势、Likes/Comments/Shares、样品寄出到内容产出的转化率

运营数据和达人组数据都拉了的话，「两段对比·归因」页面最下面还有一块**达人侧同期对照**：同一本期/基期窗口换到达人组自己的数据里看供给端（内容量/参与达人数/内容质量）跟渠道 GMV 涨跌方向对不对得上，帮着判断是达人供给变化还是达人组以外的因素在起作用。

配色主题在 `.streamlit/config.toml`（深色系 + 品牌粉红色 `#C2416B` 强调色，跟图表里"达人"渠道的颜色是同一个色，界面和图表统一视觉）。想换颜色就改这个文件的 `primaryColor` 等几行，不用碰代码。
