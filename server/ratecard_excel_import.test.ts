import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import * as XLSX from 'xlsx';
import fs from 'fs/promises';
import fsSync from 'fs';
import path from 'path';

// 系统的 10 项全局标准内置预制列
const PRESET_COLUMNS = [
  '项目组',
  '项目名称',
  '单位',
  '数量',
  '不含税单价',
  '不含税总价',
  '税率',
  '含税单价',
  '含税总价',
  '说明'
] as const;

describe('协议定价表 Excel 导入端到端测试（组名合并单元格解包与标准列映射转换）', () => {
  const testDir = path.join(process.cwd(), 'data', 'test_tmp');
  const excelFilePath = path.join(testDir, '带合并单元格示例.xlsx');
  const targetJsonPath = path.join(process.cwd(), 'data', 'ratecard', '测试端到端内置列定价表.json');

  // 包含合并单元格的数据（项目组 A2:A3 合并，A4:A5 合并）
  const sampleExcelData = [
    {
      '项目组': '电缆类',
      '物料名称': '铜芯电力电缆',
      '规格型号与描述': 'YJV-3*25+1*16 国标纯铜',
      '计量单位': '米',
      '协议不含税单价(元)': '125.50',
      '适用税率': '0.13'
    },
    {
      '项目组': '', // 跨行合并留空
      '物料名称': '铝芯电力电缆',
      '规格型号与描述': 'YJLV-4*120 国标铝芯',
      '计量单位': '米',
      '协议不含税单价(元)': '45.00',
      '适用税率': '0.13'
    },
    {
      '项目组': '紧固件类',
      '物料名称': '不锈钢螺栓',
      '规格型号与描述': 'M12*45 304不锈钢外六角',
      '计量单位': '个',
      '协议不含税单价(元)': '3.20',
      '适用税率': '0.13'
    },
    {
      '项目组': '', // 跨行合并留空
      '物料名称': '镀锌螺母',
      '规格型号与描述': 'M12 8级高强度',
      '计量单位': '个',
      '协议不含税单价(元)': '0.80',
      '适用税率': '0.13'
    },
    {
      '项目组': '仪器仪表类',
      '物料名称': '数字万用表',
      '规格型号与描述': 'FLUKE 15B+ 高精度数字万用表',
      '计量单位': '台',
      '协议不含税单价(元)': '480.00',
      '适用税率': '0.13'
    }
  ];

  beforeAll(async () => {
    await fs.mkdir(testDir, { recursive: true });
    await fs.mkdir(path.dirname(targetJsonPath), { recursive: true });
    
    const worksheet = XLSX.utils.json_to_sheet(sampleExcelData);
    // 配置合并单元格
    worksheet['!merges'] = [
      { s: { r: 1, c: 0 }, e: { r: 2, c: 0 } },
      { s: { r: 3, c: 0 }, e: { r: 4, c: 0 } },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '协议定价表');
    
    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    await fs.writeFile(excelFilePath, buffer);
  });

  afterAll(async () => {
    if (fsSync.existsSync(excelFilePath)) {
      await fs.unlink(excelFilePath).catch(() => {});
    }
    if (fsSync.existsSync(testDir)) {
      await fs.rmdir(testDir).catch(() => {});
    }
    if (fsSync.existsSync(targetJsonPath)) {
      await fs.unlink(targetJsonPath).catch(() => {});
    }
  });

  it('1. 端到端：解析 Excel -> 自动解包合并单元格组名向下继承 -> 提取行数据', async () => {
    const fileBuffer = await fs.readFile(excelFilePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // 解包合并单元格
    if (sheet && sheet['!merges']) {
      sheet['!merges'].forEach((range: any) => {
        const startCell = sheet[XLSX.utils.encode_cell(range.s)];
        if (!startCell) return;
        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            sheet[XLSX.utils.encode_cell({ r: R, c: C })] = { ...startCell };
          }
        }
      });
    }

    const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const rawHeaders: string[] = (jsonData[0] || []).map(h => String(h || '').trim());
    const headers = rawHeaders.filter(h => h.length > 0);

    const allRows: Record<string, string>[] = [];
    const lastGroupValues: Record<number, string> = {};

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || row.length === 0) continue;
      const item: Record<string, string> = {};
      let hasVal = false;

      headers.forEach((h, colIdx) => {
        let val = String(row[colIdx] ?? '').trim();
        if (!val && colIdx < 3 && lastGroupValues[colIdx]) {
          val = lastGroupValues[colIdx];
        }
        if (val) {
          hasVal = true;
          lastGroupValues[colIdx] = val;
        }
        item[h] = val;
      });

      if (hasVal) {
        allRows.push(item);
      }
    }

    expect(allRows).toHaveLength(5);
    // 验证合并单元格的组名成功继承解包
    expect(allRows[0]['项目组']).toBe('电缆类');
    expect(allRows[1]['项目组']).toBe('电缆类');
    expect(allRows[2]['项目组']).toBe('紧固件类');
    expect(allRows[3]['项目组']).toBe('紧固件类');
  });

  it('2. 端到端：用户选择列映射配置 -> 数据归一化为纯标准列 -> 保存只包含 columns 和 items 的干净 JSON', async () => {
    const fileBuffer = await fs.readFile(excelFilePath);
    const workbook = XLSX.read(fileBuffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    if (sheet && sheet['!merges']) {
      sheet['!merges'].forEach((range: any) => {
        const startCell = sheet[XLSX.utils.encode_cell(range.s)];
        if (!startCell) return;
        for (let R = range.s.r; R <= range.e.r; ++R) {
          for (let C = range.s.c; C <= range.e.c; ++C) {
            sheet[XLSX.utils.encode_cell({ r: R, c: C })] = { ...startCell };
          }
        }
      });
    }

    const jsonData: any[][] = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const headers: string[] = (jsonData[0] || []).map(h => String(h || '').trim());

    const items: Record<string, string>[] = [];
    const lastGroupValues: Record<number, string> = {};

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i];
      if (!row || row.length === 0) continue;
      const item: Record<string, string> = {};
      headers.forEach((h, colIdx) => {
        let val = String(row[colIdx] ?? '').trim();
        if (!val && colIdx < 3 && lastGroupValues[colIdx]) {
          val = lastGroupValues[colIdx];
        }
        if (val) lastGroupValues[colIdx] = val;
        item[h] = val;
      });
      items.push(item);
    }

    // 用户在界面上确定的标准列映射配置
    const mapping: Record<string, string> = {
      '项目组': '项目组',
      '物料名称': '项目名称',
      '规格型号与描述': '说明',
      '计量单位': '单位',
      '协议不含税单价(元)': '不含税单价',
      '适用税率': '税率'
    };

    const usedPresetCols = Array.from(new Set(Object.values(mapping).filter(v => Boolean(v) && PRESET_COLUMNS.includes(v as any))));
    const targetColumns = PRESET_COLUMNS.filter(col => usedPresetCols.includes(col));

    const mappedItems: Record<string, string>[] = [];
    for (const rawItem of items) {
      const newItem: Record<string, string> = {};
      let hasValue = false;
      for (const [rawHeader, presetCol] of Object.entries(mapping)) {
        if (presetCol && typeof presetCol === 'string' && PRESET_COLUMNS.includes(presetCol as any)) {
          const val = String(rawItem[rawHeader] ?? '').trim();
          if (val) hasValue = true;
          newItem[presetCol] = val;
        }
      }
      if (hasValue) {
        mappedItems.push(newItem);
      }
    }

    const payload = {
      columns: targetColumns,
      items: mappedItems
    };

    await fs.writeFile(targetJsonPath, JSON.stringify(payload, null, 2), 'utf-8');

    // 验证保存的 JSON 结构与字段极其纯净规范
    const savedData = JSON.parse(await fs.readFile(targetJsonPath, 'utf-8'));
    expect(Object.keys(savedData)).toEqual(['columns', 'items']);
    expect(savedData.columns).toEqual(['项目组', '项目名称', '单位', '不含税单价', '税率', '说明']);
    expect(savedData.items[1]).toEqual({
      '项目组': '电缆类',
      '项目名称': '铝芯电力电缆',
      '说明': 'YJLV-4*120 国标铝芯',
      '单位': '米',
      '不含税单价': '45.00',
      '税率': '0.13'
    });
  });
});
