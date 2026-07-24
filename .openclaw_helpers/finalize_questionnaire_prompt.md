[STRICT WORKSPACE RULES]
- Do not read, edit, overwrite, or generate these paths: fix_docx_format.py, openclaw.sh, logs, result, .openclaw_helpers.
- Write survey outputs only in the current working directory, not in subdirectories.
- Preserve the 47-question source form exactly: final DOCX must include question numbers 1 through 47, without omissions, merged questions, renumbering, or a replacement survey schema.
- Final DOCX must retain the original source questionnaire text for every question: include the original question stem and the original option text before the answer/evidence. Do not output only "Q1/Q2" plus answers.
- Use this per-question structure in the DOCX/Markdown draft whenever possible: `Q<number>` -> `【原题】<original question text>` -> `【原选项】<all original option text, if any>` -> `【答案】...` -> `【证据】...` -> `【置信度】...`.
- The JSON must be a flat object with exactly these string keys: 2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46.
- For numeric score questions, values must be JSON numbers only: A=1, B=2, C=3, D=4. Do not use letters and do not quote numeric values. Numeric-score-only keys are: 3,4,5,6,7,8,9,10,11,12,13,14,16,18,19,20,21,22,23,24,26,28,29,30,31,32,33,34,37.
- For Q2（主要研究领域）, do not calculate or save a score. Save selected research-field names exactly as text, for example `{"selected":["生命科学","信息科技"]}`.
- For selected-option-only questions other than Q2, do not calculate or save a score. Save only selected option letters as an object like `{"selected":["D","E","F"]}`. Selected-option-only keys are: 17,25,35,36,39,40,41,42,43,44,45,46. In particular, Q17（AI主要应用环节）and Q25（科研组织机制调整内容）are not scored; they only record selected options.
- For Q15（AI支撑平台使用情况）, save both selected options and score as `{"selected":["B","G"],"score":3}`. Scoring rule: if A is selected, score 1; if A is not selected and exactly 1 other option is selected, score 2; if 2 to 4 non-A options are selected, score 3; if more than 4 non-A options are selected, score 4.
- For Q27（分布式科研团队组建情况）, save both selected options and score as `{"selected":["B","C"],"score":3}`. Scoring rule: selected A means 1; if selected options include B, score at least 2; if they include C, score at least 3; if they include D, score 4. Use the highest matching score among A/B/C/D; do not use the generic multi-choice count rule for Q27.
- For Q38（改革紧迫性）, save the selected option and score as `{"selected":["C"],"score":3}`. Q38 is the only demand-section question with a score.
- For Q39-Q46, do not calculate or save scores. Save only selected option letters as `{"selected":[...]}`.
- Do not save option text, evidence, confidence, or explanations in JSON.
- For question 47, rank 2 to 4 demands using exactly this format: 第一：<dimension> - <specific demand>. Each <dimension> must be one of: 建立扁平化组织架构；建立跨学科分布式团队；改革考核评价机制；明确知识产权归属；优化资源配置；引进与培养复合型人才；培育开放合作与协同文化. Do not add a summary line such as `【答案】第一：...` before the ranked list. Each rank label（第一/第二/第三/第四）may appear at most once in Q47, and ranked dimensions must not repeat.

【输出目录硬性要求】
- 不要新建机构缩写目录或任何子目录保存结果。
- 所有生成文件必须直接写入当前工作目录：{{CURRENT_DIR}}。
- 读取 ./问卷.md 和 ./lost.md；最终 Word 必须是 ./{{RAW_DOCX}} 或 ./问卷v2.docx；单选答案必须是 ./{{SINGLE_CHOICE_JSON}}。

你可以使用 skills 文件夹下的技能。你的任务是检索 {{INSTITUTE_NAME}} 的信息来补全答卷问题，原始答卷文件为《{{SOURCE_QUESTIONNAIRE_DOCX}}》。

最终《问卷v2.md》和 Word 必须保留源问卷每一题的原始题干和原始选项文字。每题建议使用以下结构：Q题号、【原题】、【原选项】、【答案】、【证据】、【置信度】。不要只输出 Q1/Q2 和答案，否则最终材料无法脱离源问卷审阅。

置信度分为 [100,80,60,40,20,0] 6 部分，分别对应 [绝对事实，有可靠证据支撑，不一定，非常怀疑，缺乏相关证据，没有任何证据支撑]。输出的置信度打分必须落到这几个离散数值中。

你现在需要：
1. 对在 问卷.md 中置信度小于 80 的问题进行全网检索，结合 lost.md 中指出的问题，对《问卷.md》进行针对性修改，生成《问卷v2.md》文件。其中第三个问题，机构的研究领域部分，每个研究领域必须要有证据支撑你的观点。
2. 确保你的输出内容置信度至少 80 并给出相关证据支撑。注意置信度要准确真实，如果实在无法满足置信度要求也要照实输出。如果问题回答的置信度依赖进一步检索网页和 arxiv，可进一步扩大检索。
3. 最后尝试将《问卷v2.md》文件转换为《{{RAW_DOCX}}》文件。
4. 同时生成《{{SINGLE_CHOICE_JSON}}》文件。JSON 必须是一个扁平对象，键必须正好是字符串题号 "2" 到 "46"。不要保存填空题、开放题、选项文字、证据、置信度或解释。字段格式必须严格遵守本提示开头的 JSON 规则：Q2 保存研究领域名称列表；纯计分题保存数字；只统计选项的题保存 `{"selected":[...]}`；Q15、Q27、Q38 保存 `{"selected":[...],"score":数字}`。示例：{"2": {"selected": ["生命科学"]}, "3": 2, "15": {"selected": ["B", "G"], "score": 3}, "17": {"selected": ["D", "E", "F"]}, "27": {"selected": ["B", "C"], "score": 3}, "38": {"selected": ["C"], "score": 3}, "39": {"selected": ["B", "D"]}}。

在进行任何文件读取或编辑之前，你必须先使用工具查看当前目录下的文件列表（如执行 ls 或类似技能），确认文件存在且名称无误后，再进行操作。注意最终问卷文件要包括原始问题和回答。多选问题所选的每个选项都要给出确切的证据，不得所有选项都选择并且泛泛给出依据。单选计分题答案 JSON 必须按选项顺序映射为数字：A=1、B=2、C=3、D=4。只统计选项的多选题 JSON 必须保存 selected 列表，不要保存 score；Q15、Q27、Q38 必须同时保存 selected 和 score。
