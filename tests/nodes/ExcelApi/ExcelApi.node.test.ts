import { ExcelApi } from '../../../nodes/ExcelApi/ExcelApi.node';
import { MockExecuteFunctions } from '../../helpers/mockHelpers';
import { NodeOperationError } from 'n8n-workflow';

describe('ExcelApi Node', () => {
  let excelApi: ExcelApi;
  let mockFunctions: MockExecuteFunctions;

  beforeEach(() => {
    excelApi = new ExcelApi();
    mockFunctions = new MockExecuteFunctions();
    
    // 設定預設憑證
    mockFunctions.setCredentials('excelApiAuth', {
      url: 'http://localhost:8000',
      token: 'test-token-123',
    });
  });

  describe('Node Properties', () => {
    it('should have correct node properties', () => {
      expect(excelApi.description.displayName).toBe('Excel API');
      expect(excelApi.description.name).toBe('excelApi');
      expect(excelApi.description.group).toContain('transform');
      expect(excelApi.description.version).toBe(1);
    });

    it('should require credentials', () => {
      const credentials = excelApi.description.credentials;
      expect(credentials).toBeDefined();
      expect(credentials?.[0].name).toBe('excelApiAuth');
      expect(credentials?.[0].required).toBe(true);
    });

    it('should have all operations defined', () => {
      const operationProperty = excelApi.description.properties.find(
        (p) => p.name === 'operation'
      );
      expect(operationProperty).toBeDefined();
      expect(operationProperty?.options).toHaveLength(5);
      
      const operations = operationProperty?.options as any[];
      expect(operations.map(o => o.value)).toEqual([
        'append', 'read', 'update', 'delete', 'batch'
      ]);
    });
  });

  describe('Load Options Methods', () => {
    describe('getExcelFiles', () => {
      it('should load Excel files successfully', async () => {
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/files',
          {
            success: true,
            files: ['employees.xlsx', 'sales.xlsx', 'inventory.xlsx'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const files = await excelApi.methods.loadOptions.getExcelFiles.call(
          loadFunctions
        );

        expect(files).toHaveLength(3);
        expect(files[0]).toEqual({ name: 'employees.xlsx', value: 'employees.xlsx' });
        expect(files[1]).toEqual({ name: 'sales.xlsx', value: 'sales.xlsx' });
      });

      it('should return empty array on API error', async () => {
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/files',
          { error: new Error('API Error') }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const files = await excelApi.methods.loadOptions.getExcelFiles.call(
          loadFunctions
        );

        expect(files).toEqual([]);
      });
    });

    describe('getExcelSheets', () => {
      it('should load sheets from selected file', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/sheets?file=employees.xlsx',
          {
            success: true,
            sheets: ['Sheet1', 'Sheet2', 'Summary'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const sheets = await excelApi.methods.loadOptions.getExcelSheets.call(
          loadFunctions
        );

        expect(sheets).toHaveLength(3);
        expect(sheets[0]).toEqual({ name: 'Sheet1', value: 'Sheet1' });
      });

      it('should return empty array when no file selected', async () => {
        mockFunctions.setParameter('fileName', '');

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const sheets = await excelApi.methods.loadOptions.getExcelSheets.call(
          loadFunctions
        );

        expect(sheets).toEqual([]);
      });
    });
  });

  describe('Execute Method - Append Operation', () => {
    beforeEach(() => {
      mockFunctions.setParameter('operation', 'append', 0);
      mockFunctions.setParameter('fileName', 'employees.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);
    });

    it('should append row in object mode', async () => {
      mockFunctions.setParameter('appendMode', 'object', 0);
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        '員工編號': 'E100',
        '姓名': '測試員工',
        '部門': '技術部',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 5,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result).toHaveLength(1);
      expect(result[0]).toHaveLength(1);
      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.row).toBe(5);
    });

    it('should append row in array mode', async () => {
      mockFunctions.setParameter('appendMode', 'array', 0);
      mockFunctions.setParameter('appendValuesArray', JSON.stringify([
        'E100', '測試員工', '技術部'
      ]), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append',
        {
          success: true,
          message: 'Row appended successfully',
          row: 5,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result).toHaveLength(1);
      expect(result[0][0].json.success).toBe(true);
    });

    it('should throw error for invalid JSON in object mode', async () => {
      mockFunctions.setParameter('appendMode', 'object', 0);
      mockFunctions.setParameter('appendValuesObject', 'invalid json', 0);

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow(NodeOperationError);
    });
  });

  describe('Execute Method - Read Operation', () => {
    beforeEach(() => {
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('fileName', 'employees.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);
    });

    it('should read data with headers and convert to objects', async () => {
      mockFunctions.setParameter('range', '', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          success: true,
          data: [
            ['員工編號', '姓名', '部門'],
            ['E001', '張三', '技術部'],
            ['E002', '李四', '人資部'],
          ],
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result).toHaveLength(1);
      expect(result[0]).toHaveLength(2);
      expect(result[0][0].json).toEqual({
        '員工編號': 'E001',
        '姓名': '張三',
        '部門': '技術部',
      });
      expect(result[0][1].json).toEqual({
        '員工編號': 'E002',
        '姓名': '李四',
        '部門': '人資部',
      });
    });

    it('should read data with range specified', async () => {
      mockFunctions.setParameter('range', 'A1:C3', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          success: true,
          data: [
            ['A1', 'B1', 'C1'],
            ['A2', 'B2', 'C2'],
          ],
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result).toHaveLength(1);
      expect(result[0]).toHaveLength(1);
    });
  });

  describe('Execute Method - Update Operation', () => {
    beforeEach(() => {
      mockFunctions.setParameter('operation', 'update', 0);
      mockFunctions.setParameter('fileName', 'employees.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);
    });

    it('should update row by row number', async () => {
      mockFunctions.setParameter('identifyBy', 'rowNumber', 0);
      mockFunctions.setParameter('rowNumber', 5, 0);
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        '狀態': '已完成',
        '更新日期': '2025-12-21',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          message: 'Row updated successfully',
          row: 5,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.row).toBe(5);
    });

    it('should update row by lookup', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E100', 0);
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        '薪資': '80000',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          message: 'Row updated successfully',
          row: 5,
          matched_value: 'E100',
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.matched_value).toBe('E100');
    });
  });

  describe('Execute Method - Delete Operation', () => {
    beforeEach(() => {
      mockFunctions.setParameter('operation', 'delete', 0);
      mockFunctions.setParameter('fileName', 'employees.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);
    });

    it('should delete row by row number', async () => {
      mockFunctions.setParameter('identifyBy', 'rowNumber', 0);
      mockFunctions.setParameter('rowNumber', 5, 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          message: 'Row deleted successfully',
          row: 5,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
    });

    it('should delete row by lookup', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E100', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          message: 'Row deleted successfully',
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
    });
  });

  describe('Execute Method - Batch Operation', () => {
    it('should execute batch operations', async () => {
      mockFunctions.setParameter('operation', 'batch', 0);
      mockFunctions.setParameter('fileName', 'employees.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setParameter('batchOperations', JSON.stringify([
        { type: 'append', values: ['E100', '測試'] },
        { type: 'update', row: 5, values: ['E005', '更新'] },
      ]), 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/batch',
        {
          success: true,
          results: [
            { success: true, operation: 'append' },
            { success: true, operation: 'update' },
          ],
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.results).toHaveLength(2);
    });
  });

  describe('Error Handling', () => {
    it('should throw error when file name is missing', async () => {
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('fileName', '', 0);
      mockFunctions.setInputData([{ json: {} }]);

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow('File Name is required');
    });

    it('should handle API errors gracefully', async () => {
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('fileName', 'test.xlsx', 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          error: {
            response: {
              body: { error: 'File not found' },
            },
          },
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow();
    });
  });
});
