"""Add the 2 missing single choice explanations to the merged file."""
import json

# Load the current merged file
with open('_questions_with_explanations.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Add explanation for ID 211
# Question: 成文日期通常是指（ ）。
# Answer: D (发文机关负责人签发的日期)
for q in data['singles']:
    if q['id'] == 211:
        q['explanation'] = (
            '成文日期是公文生效的法定时间。根据《党政机关公文处理工作条例》规定，'
            '公文一般以机关负责人签发的日期为准。选项A草拟公文文稿的日期属于起草阶段，'
            '选项B公文印制完毕的日期属于印制环节，选项C领导人在公文正本上签署的日期是签署行为，'
            '均不是成文日期的准确含义。成文日期的核心是发文机关负责人签发的日期，'
            '是公文开始具有法定效力的时间节点。'
        )

    if q['id'] == 212:
        q['explanation'] = (
            '公文行文规范包括三个方面的内容：行文关系、行文方向与方式、行文规则。'
            '行文关系根据隶属关系和职权范围确定；行文方向分为上行、下行和平行，'
            '行文方式包括逐级行文、多级行文、越级行文等；行文规则是对行文过程中的规范性要求。'
            '选项A缺少行文规则，选项B缺少行文关系，选项C的表述最完整准确。'
        )

# Count
total_q = sum(len(v) for v in data.values())
total_exp = sum(1 for v in data.values() for q in v if q.get('explanation'))
print(f'Updated: {total_exp}/{total_q} questions have explanations.')

# Save
with open('_questions_with_explanations.json', 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print('Saved to _questions_with_explanations.json')
