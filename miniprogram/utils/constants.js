// 10,000 字符不足以容纳一篇含公式的论文；100,000 仍低于云函数的单次请求容量。
const MAX_SOURCE_LENGTH = 100000;
const TEMPLATES = [
  { id: 'homework', name: '课程作业', description: '适合日常作业与课程报告', cover: false },
  { id: 'lab-report', name: '实验报告', description: '带封面和固定章节结构', cover: true, requiredSections: ['实验目的', '实验原理', '实验步骤', '实验结果', '结论'] },
  { id: 'study-note', name: '学术笔记', description: '适合整理阅读与学习内容', cover: false }
];
module.exports = { MAX_SOURCE_LENGTH, TEMPLATES };
