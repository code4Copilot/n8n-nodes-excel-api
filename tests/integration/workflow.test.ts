import { ExcelApi } from '../../nodes/ExcelApi/ExcelApi.node';
import { MockExecuteFunctions } from '../helpers/mockHelpers';

describe('Excel API Integration Tests', () => {
  let excelApi: ExcelApi;
  let mockFunctions: MockExecuteFunctions;

  beforeEach(() => {
    excelApi = new ExcelApi();
    mockFunctions = new MockExecuteFunctions();
    
    mockFunctions.setCredentials('excelApiAuth', {
      url: 'http://localhost:8000',
      token: 'test-token',
    });
  });

  describe('Complete Workflow: Append → Read → Update → Delete', () => {
    it('should handle complete workflow', async () => {
      // 1. Append
      mockFunctions.setParameter('operation', 'append', 0);
      mockFunctions.setParameter('fileName', 'test.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setParameter('appendMode', 'object', 0);
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'ID': 'TEST001',
        'Name': 'Test User',
      }), 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        { success: true, row: 2 }
      );

      let executeFunctions = mockFunctions.getExecuteFunctions();
      let result = await excelApi.execute.call(executeFunctions);
      expect(result[0][0].json.success).toBe(true);

      // 2. Read
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('range', '', 0);
      
      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          success: true,
          data: [
            ['ID', 'Name'],
            ['TEST001', 'Test User'],
          ],
        }
      );

      executeFunctions = mockFunctions.getExecuteFunctions();
      result = await excelApi.execute.call(executeFunctions);
      expect(result[0][0].json.ID).toBe('TEST001');

      // 3. Update
      mockFunctions.setParameter('operation', 'update', 0);
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', 'ID', 0);
      mockFunctions.setParameter('lookupValue', 'TEST001', 0);
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        'Name': 'Updated User',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        { success: true, row: 2 }
      );

      executeFunctions = mockFunctions.getExecuteFunctions();
      result = await excelApi.execute.call(executeFunctions);
      expect(result[0][0].json.success).toBe(true);

      // 4. Delete
      mockFunctions.setParameter('operation', 'delete', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        { success: true }
      );

      executeFunctions = mockFunctions.getExecuteFunctions();
      result = await excelApi.execute.call(executeFunctions);
      expect(result[0][0].json.success).toBe(true);
    });
  });

  describe('Batch Operations', () => {
    it('should handle multiple operations in batch', async () => {
      mockFunctions.setParameter('operation', 'batch', 0);
      mockFunctions.setParameter('fileName', 'test.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setParameter('batchOperations', JSON.stringify([
        { type: 'append', values: ['E100', '員工A', '技術部'] },
        { type: 'append', values: ['E101', '員工B', '人資部'] },
        { type: 'update', row: 3, values: ['E100', '更新員工', '技術部'] },
      ]), 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/batch',
        {
          success: true,
          results: [
            { success: true, operation: 'append', row: 5 },
            { success: true, operation: 'append', row: 6 },
            { success: true, operation: 'update', row: 3 },
          ],
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.results).toHaveLength(3);
    });
  });

  describe('Error Recovery', () => {
    it('should handle API errors during workflow', async () => {
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('fileName', 'nonexistent.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          error: {
            response: {
              body: { error: 'File not found: nonexistent.xlsx' },
              statusCode: 404,
            },
          },
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow();
    });

    it('should handle network timeout errors', async () => {
      mockFunctions.setParameter('operation', 'append', 0);
      mockFunctions.setParameter('fileName', 'test.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setParameter('appendMode', 'array', 0);
      mockFunctions.setParameter('appendValuesArray', JSON.stringify(['data']), 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append',
        {
          error: new Error('ETIMEDOUT: Request timeout'),
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow();
    });
  });

  describe('Data Validation', () => {
    it('should handle empty data sets', async () => {
      mockFunctions.setParameter('operation', 'read', 0);
      mockFunctions.setParameter('fileName', 'empty.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/read',
        {
          success: true,
          data: [],
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result).toHaveLength(1);
      expect(result[0]).toHaveLength(1);
    });

    it('should handle special characters in data', async () => {
      mockFunctions.setParameter('operation', 'append', 0);
      mockFunctions.setParameter('fileName', 'test.xlsx', 0);
      mockFunctions.setParameter('sheetName', 'Sheet1', 0);
      mockFunctions.setParameter('appendMode', 'object', 0);
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        '名稱': '測試員工 & Co. <特殊>',
        '說明': '包含 "引號" 和 \'單引號\'',
      }), 0);
      mockFunctions.setInputData([{ json: {} }]);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          row: 3,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
    });
  });
});
