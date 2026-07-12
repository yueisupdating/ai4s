[STRICT WORKSPACE RULES]
- Do not read, edit, overwrite, or generate these paths: fix_docx_format.py, openclaw.sh, logs, result, .openclaw_helpers.
- Write survey outputs only in the current working directory, not in subdirectories.
- Preserve the 47-question source form exactly: final DOCX must include question numbers 1 through 47, without omissions, merged questions, renumbering, or a replacement survey schema.
- Final DOCX must retain the original source questionnaire text for every question: include the original question stem and the original option text before the answer/evidence. Do not output only "Q1/Q2" plus answers.
- Use this per-question structure in the DOCX/Markdown draft whenever possible: `Q<number>` -> `【原题】<original question text>` -> `【原选项】<all original option text, if any>` -> `【答案】...` -> `【证据】...` -> `【置信度】...`.
- The JSON must be a flat object with exactly these string keys: 2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46.
- Single-choice values must be JSON numbers only: A=1, B=2, C=3, D=4. Do not use letters and do not quote numeric values.
- Multi-choice values must be JSON score numbers only: if the first option is selected, score 1; if the first option is not selected and exactly 1 other option is selected, score 2; if 2 to 4 non-first options are selected, score 3; if more than 4 non-first options are selected, score 4.
- For question 47, rank 2 to 4 demands using exactly this format: 第一：<dimension> - <specific demand>. Each <dimension> must be one of: 建立扁平化组织架构；建立跨学科分布式团队；改革考核评价机制；明确知识产权归属；优化资源配置；引进与培养复合型人才；培育开放合作与协同文化.

【输出目录硬性要求】
- 不要新建机构缩写目录或任何子目录保存结果。
- 所有生成文件必须直接写入当前工作目录：{{CURRENT_DIR}}。
- 规划文件必须是 ./plan.md；中间答卷必须是 ./问卷.md、./lost.md、./问卷v2.md；最终 Word 必须是 ./{{RAW_DOCX}} 或 ./问卷v2.docx；单选答案必须是 ./{{SINGLE_CHOICE_JSON}}。

你可以使用 skills 文件夹下的技能。你需要整合 {{INSTITUTE_NAME}} 的信息，用于回答《{{SOURCE_QUESTIONNAIRE_DOCX}}》的答卷。我目前打算先让你先将答卷文件转为 问卷.md 文件。

生成《问卷.md》和最终 Word 时，必须保留源问卷每一题的原始题干和原始选项文字。每题建议使用以下结构：Q题号、【原题】、【原选项】、【答案】、【证据】、【置信度】。不要只输出 Q1/Q2 和答案，否则最终材料无法脱离源问卷审阅。

置信度分为 [100,80,60,40,20,0] 6 部分，分别对应 [绝对事实，有可靠证据支撑，不一定，非常怀疑，缺乏相关证据，没有任何证据支撑]。输出的置信度打分必须落到这几个离散数值中。

在查询信息时，为保证信息的置信度，你只查询 {{INSTITUTE_NAME}} 相关官方网站以及目录下收集的 arxiv.md 文件，将相关信息持久化到 md 文件里。整合信息，填写 问卷.md 文件，问卷.md 文件内容置信度必须超过 80，将问题的信息来源 URL 明确指出。并将缺失信息的问题列出来，持久化到 lost.md 里，分析缺失信息需要从哪些网页或渠道检索，或者当前方案的不足之处，分条写在 lost.md 文件里。

请先在当前目录下生成规划文件，文件名必须严格命名为 plan.md。第二步：读取 plan.md 的内容再执行后续步骤。在进行任何文件读取或编辑之前，你必须先使用工具查看当前目录下的文件列表（如执行 ls 或类似技能），确认文件存在且名称无误后，再进行操作。

多选问题每个选项都要给出确切的证据，不得所有选项都选择并且泛泛给出依据。单选题按选项顺序映射为数字：A=1、B=2、C=3、D=4，回答时只给出数字 1、2、3、4 中的一项而不用复述选项文字。多选题用于 JSON 时只保存得分：选第一个选项记 1 分；未选第一个且只选其他 1 项记 2 分；未选第一个且选其他 2-4 项记 3 分；未选第一个且选其他 4 项以上记 4 分。
