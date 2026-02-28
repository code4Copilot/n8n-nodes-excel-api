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
      expect(operationProperty?.options).toHaveLength(4);
      
      const operations = operationProperty?.options as any[];
      expect(operations.map(o => o.value)).toEqual([
        'append', 'read', 'update', 'delete'
      ]);
    });

    describe('reloadOnChange for dynamic dropdown fields', () => {
      it('sheetName should be type resourceLocator (always loads fresh, no caching)', () => {
        const sheetNameProp = excelApi.description.properties.find(
          (p) => p.name === 'sheetName'
        );
        expect(sheetNameProp).toBeDefined();
        expect(sheetNameProp?.type).toBe('resourceLocator');
      });

      it('sheetName list mode should use searchListMethod: searchExcelSheets', () => {
        const sheetNameProp = excelApi.description.properties.find(
          (p) => p.name === 'sheetName'
        );
        const listMode = (sheetNameProp as any)?.modes?.find((m: any) => m.type === 'list');
        expect(listMode).toBeDefined();
        expect(listMode?.typeOptions?.searchListMethod).toBe('searchExcelSheets');
      });

      it('sheetName should also have a manual string input mode as fallback', () => {
        const sheetNameProp = excelApi.description.properties.find(
          (p) => p.name === 'sheetName'
        );
        const stringMode = (sheetNameProp as any)?.modes?.find((m: any) => m.type === 'string');
        expect(stringMode).toBeDefined();
      });

      it('lookupColumn should be type resourceLocator (always loads fresh, no caching)', () => {
        const lookupColumnProp = excelApi.description.properties.find(
          (p) => p.name === 'lookupColumn'
        );
        expect(lookupColumnProp).toBeDefined();
        expect(lookupColumnProp?.type).toBe('resourceLocator');
      });

      it('lookupColumn list mode should use searchListMethod: searchColumnNames', () => {
        const lookupColumnProp = excelApi.description.properties.find(
          (p) => p.name === 'lookupColumn'
        );
        const listMode = (lookupColumnProp as any)?.modes?.find((m: any) => m.type === 'list');
        expect(listMode).toBeDefined();
        expect(listMode?.typeOptions?.searchListMethod).toBe('searchColumnNames');
      });

      it('lookupColumn should also have a manual string input mode as fallback', () => {
        const lookupColumnProp = excelApi.description.properties.find(
          (p) => p.name === 'lookupColumn'
        );
        const stringMode = (lookupColumnProp as any)?.modes?.find((m: any) => m.type === 'string');
        expect(stringMode).toBeDefined();
      });

      it('fileName should be type resourceLocator (loads fresh on every open)', () => {
        const fileNameProp = excelApi.description.properties.find(
          (p) => p.name === 'fileName'
        );
        expect(fileNameProp?.type).toBe('resourceLocator');
        const listMode = (fileNameProp as any)?.modes?.find((m: any) => m.type === 'list');
        expect(listMode?.typeOptions?.searchListMethod).toBe('searchExcelFiles');
      });
    });
  });

  describe('Load Options Methods', () => {
    describe('getExcelSheets (loadOptions fallback)', () => {
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

    describe('getColumnNames', () => {
      it('should load column names from selected file and sheet', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          {
            success: true,
            headers: ['員工編號', '姓名', '部門', '薪資', '入職日期'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(5);
        expect(columns[0]).toEqual({ name: '員工編號', value: '員工編號' });
        expect(columns[1]).toEqual({ name: '姓名', value: '姓名' });
        expect(columns[2]).toEqual({ name: '部門', value: '部門' });
      });

      it('should return empty array when no file selected', async () => {
        mockFunctions.setParameter('fileName', '');
        mockFunctions.setParameter('sheetName', 'Sheet1');

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toEqual([]);
      });

      it('should return empty array when no sheet selected', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', '');

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toEqual([]);
      });

      it('should return empty array on API error', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          { error: new Error('API Error') }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toEqual([]);
      });

      it('should handle special characters in file and sheet names', async () => {
        mockFunctions.setParameter('fileName', 'test file 測試.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet 工作表');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=test%20file%20%E6%B8%AC%E8%A9%A6.xlsx&sheet=Sheet%20%E5%B7%A5%E4%BD%9C%E8%A1%A8',
          {
            success: true,
            headers: ['Column A', 'Column B'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(2);
        expect(columns[0]).toEqual({ name: 'Column A', value: 'Column A' });
      });
    });
  });

  describe('List Search Methods (resourceLocator)', () => {
    describe('searchExcelFiles', () => {
      it('should return all files when no filter provided', async () => {
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/files',
          { success: true, files: ['employees.xlsx', 'sales.xlsx', 'inventory.xlsx'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelFiles.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(3);
        expect(result.results[0]).toEqual({ name: 'employees.xlsx', value: 'employees.xlsx' });
        expect(result.results[2]).toEqual({ name: 'inventory.xlsx', value: 'inventory.xlsx' });
      });

      it('should filter files by search text (case-insensitive)', async () => {
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/files',
          { success: true, files: ['employees.xlsx', 'Employee_backup.xlsx', 'sales.xlsx'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelFiles.call(
          loadFunctions, 'employee'
        );

        expect(result.results).toHaveLength(2);
        expect(result.results.map((r: any) => r.value)).toEqual(['employees.xlsx', 'Employee_backup.xlsx']);
      });

      it('should return empty results on API error', async () => {
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/files',
          { error: new Error('API Error') }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelFiles.call(
          loadFunctions, undefined
        );

        expect(result.results).toEqual([]);
      });
    });

    describe('searchExcelSheets', () => {
      it('should return all sheets when no filter provided', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/sheets?file=employees.xlsx',
          { success: true, sheets: ['Sheet1', 'Overview', 'Archive'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelSheets.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(3);
        expect(result.results[0]).toEqual({ name: 'Sheet1', value: 'Sheet1' });
        expect(result.results[1]).toEqual({ name: 'Overview', value: 'Overview' });
      });

      it('should filter sheets by search text (case-insensitive)', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/sheets?file=employees.xlsx',
          { success: true, sheets: ['Sheet1', 'Overview', 'Archive', 'sheet2'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelSheets.call(
          loadFunctions, 'sheet'
        );

        expect(result.results).toHaveLength(2);
        expect(result.results.map((r: any) => r.value)).toEqual(['Sheet1', 'sheet2']);
      });

      it('should return empty results when no file selected', async () => {
        mockFunctions.setParameter('fileName', '');

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelSheets.call(
          loadFunctions, undefined
        );

        expect(result.results).toEqual([]);
      });

      it('should return empty results on API error', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/sheets?file=employees.xlsx',
          { error: new Error('API Error') }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelSheets.call(
          loadFunctions, undefined
        );

        expect(result.results).toEqual([]);
      });

      it('should handle resourceLocator object format for fileName', async () => {
        mockFunctions.setParameter('fileName', { mode: 'list', value: 'employees.xlsx' });
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/sheets?file=employees.xlsx',
          { success: true, sheets: ['Sheet1', 'Data'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchExcelSheets.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(2);
        expect(result.results[0].value).toBe('Sheet1');
      });
    });

    describe('searchColumnNames', () => {
      it('should return all columns when no filter provided', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          { success: true, headers: ['ID', 'Name', 'Department', 'Salary'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchColumnNames.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(4);
        expect(result.results[0]).toEqual({ name: 'ID', value: 'ID' });
        expect(result.results[2]).toEqual({ name: 'Department', value: 'Department' });
      });

      it('should filter columns by search text (case-insensitive)', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          { success: true, headers: ['ID', 'Name', 'department_code', 'Salary'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchColumnNames.call(
          loadFunctions, 'depart'
        );

        expect(result.results).toHaveLength(1);
        expect(result.results[0].value).toBe('department_code');
      });

      it('should handle resourceLocator object format for sheetName', async () => {
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', { mode: 'list', value: 'Sheet1' });
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          { success: true, headers: ['ID', 'Name'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchColumnNames.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(2);
        expect(result.results[0].value).toBe('ID');
      });

      it('should handle resourceLocator object format for fileName', async () => {
        mockFunctions.setParameter('fileName', { mode: 'list', value: 'employees.xlsx' });
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          { success: true, headers: ['ID', 'Name'] }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchColumnNames.call(
          loadFunctions, undefined
        );

        expect(result.results).toHaveLength(2);
        expect(result.results[1].value).toBe('Name');
      });

      it('should return empty results when fileName or sheetName is missing', async () => {
        mockFunctions.setParameter('fileName', '');
        mockFunctions.setParameter('sheetName', '');

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const result = await excelApi.methods.listSearch.searchColumnNames.call(
          loadFunctions, undefined
        );

        expect(result.results).toEqual([]);
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

    it('should update row by lookup - first match only', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E100', 0);
      mockFunctions.setParameter('processMode', 'first', 0);
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
          rows_affected: 1,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.matched_value).toBe('E100');
      expect(result[0][0].json.rows_affected).toBe(1);
      
      // Verify that process_all is set to false
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(false);
    });

    it('should update all matching rows by lookup', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '部門', 0);
      mockFunctions.setParameter('lookupValue', '技術部', 0);
      mockFunctions.setParameter('processMode', 'all', 0);
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        '狀態': '已審核',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          message: 'Multiple rows updated successfully',
          rows_affected: 5,
          matched_value: '技術部',
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.rows_affected).toBe(5);
      
      // Verify that process_all is set to true
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(true);
    });

    it('should default to processMode "all" when not specified', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E100', 0);
      mockFunctions.setParameter('processMode', 'all', 0); // Default value
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        '薪資': '90000',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          message: 'Row updated successfully',
          rows_affected: 1,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(true);
    });

    it('should throw error when no matching rows found (lookup)', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E999', 0);
      mockFunctions.setParameter('processMode', 'first', 0);
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        '薪資': '90000',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          updated_count: 0, // No rows affected
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow('No matching rows found. Lookup column: "員工編號", Lookup value: "E999"');
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

    it('should delete row by lookup - first match only', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E100', 0);
      mockFunctions.setParameter('processMode', 'first', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          message: 'Row deleted successfully',
          rows_affected: 1,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.rows_affected).toBe(1);
      
      // Verify that process_all is set to false
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(false);
    });

    it('should delete all matching rows by lookup', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '狀態', 0);
      mockFunctions.setParameter('lookupValue', '已離職', 0);
      mockFunctions.setParameter('processMode', 'all', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          message: 'Multiple rows deleted successfully',
          rows_affected: 3,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      expect(result[0][0].json.rows_affected).toBe(3);
      
      // Verify that process_all is set to true
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(true);
    });

    it('should default to processMode "all" when not specified', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E999', 0);
      mockFunctions.setParameter('processMode', 'all', 0); // Default value

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          message: 'Row deleted successfully',
          rows_affected: 1,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      const result = await excelApi.execute.call(executeFunctions);

      expect(result[0][0].json.success).toBe(true);
      
      const requestBody = mockFunctions.getLastRequestBody();
      expect(requestBody.process_all).toBe(true);
    });

    it('should throw error when no matching rows found (lookup)', async () => {
      mockFunctions.setParameter('identifyBy', 'lookup', 0);
      mockFunctions.setParameter('lookupColumn', '員工編號', 0);
      mockFunctions.setParameter('lookupValue', 'E999', 0);
      mockFunctions.setParameter('processMode', 'first', 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/delete_advanced',
        {
          success: true,
          deleted_count: 0, // No rows affected
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();

      await expect(
        excelApi.execute.call(executeFunctions)
      ).rejects.toThrow('No matching rows found. Lookup column: "員工編號", Lookup value: "E999"');
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

  describe('Lookup Column Selection Feature', () => {
    describe('Integration Test - Column Selection to Execution', () => {
      it('should load columns and use selected column for update operation', async () => {
        // Step 1: Load column names
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          {
            success: true,
            headers: ['員工編號', '姓名', '部門', '薪資'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(4);
        expect(columns.map(c => c.value)).toContain('員工編號');

        // Step 2: Use the selected column in update operation
        mockFunctions.setParameter('operation', 'update', 0);
        mockFunctions.setParameter('identifyBy', 'lookup', 0);
        mockFunctions.setParameter('lookupColumn', '員工編號', 0); // Selected from dropdown
        mockFunctions.setParameter('lookupValue', 'E001', 0);
        mockFunctions.setParameter('processMode', 'first', 0);
        mockFunctions.setParameter('valuesToSet', JSON.stringify({
          '薪資': '85000',
        }), 0);
        mockFunctions.setInputData([{ json: {} }]);

        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/update_advanced',
          {
            success: true,
            message: 'Row updated successfully',
            row: 2,
            rows_affected: 1,
          }
        );

        const executeFunctions = mockFunctions.getExecuteFunctions();
        const result = await excelApi.execute.call(executeFunctions);

        expect(result[0][0].json.success).toBe(true);
        
        // Verify that the correct lookup column was sent to API
        const requestBody = mockFunctions.getLastRequestBody();
        expect(requestBody.lookup_column).toBe('員工編號');
        expect(requestBody.lookup_value).toBe('E001');
      });

      it('should load columns and use selected column for delete operation', async () => {
        // Step 1: Load column names
        mockFunctions.setParameter('fileName', 'employees.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=employees.xlsx&sheet=Sheet1',
          {
            success: true,
            headers: ['員工編號', '姓名', '狀態'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(3);
        expect(columns.map(c => c.value)).toContain('狀態');

        // Step 2: Use the selected column in delete operation
        mockFunctions.setParameter('operation', 'delete', 0);
        mockFunctions.setParameter('identifyBy', 'lookup', 0);
        mockFunctions.setParameter('lookupColumn', '狀態', 0); // Selected from dropdown
        mockFunctions.setParameter('lookupValue', '已離職', 0);
        mockFunctions.setParameter('processMode', 'all', 0);
        mockFunctions.setInputData([{ json: {} }]);

        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/delete_advanced',
          {
            success: true,
            message: 'Multiple rows deleted successfully',
            rows_affected: 3,
          }
        );

        const executeFunctions = mockFunctions.getExecuteFunctions();
        const result = await excelApi.execute.call(executeFunctions);

        expect(result[0][0].json.success).toBe(true);
        expect(result[0][0].json.rows_affected).toBe(3);
        
        // Verify that the correct lookup column was sent to API
        const requestBody = mockFunctions.getLastRequestBody();
        expect(requestBody.lookup_column).toBe('狀態');
        expect(requestBody.lookup_value).toBe('已離職');
        expect(requestBody.process_all).toBe(true);
      });

      it('should handle columns with special characters and spaces', async () => {
        mockFunctions.setParameter('fileName', 'test.xlsx');
        mockFunctions.setParameter('sheetName', 'Sheet1');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=test.xlsx&sheet=Sheet1',
          {
            success: true,
            headers: ['Employee ID', 'Full Name (EN)', '電子郵件地址', 'Department/部門'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(4);
        expect(columns[1].value).toBe('Full Name (EN)');
        expect(columns[3].value).toBe('Department/部門');

        // Use column with special characters
        mockFunctions.setParameter('operation', 'update', 0);
        mockFunctions.setParameter('identifyBy', 'lookup', 0);
        mockFunctions.setParameter('lookupColumn', 'Full Name (EN)', 0);
        mockFunctions.setParameter('lookupValue', 'John Doe', 0);
        mockFunctions.setParameter('processMode', 'first', 0);
        mockFunctions.setParameter('valuesToSet', JSON.stringify({
          'Department/部門': 'IT',
        }), 0);
        mockFunctions.setInputData([{ json: {} }]);

        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/update_advanced',
          {
            success: true,
            message: 'Row updated successfully',
            rows_affected: 1,
          }
        );

        const executeFunctions = mockFunctions.getExecuteFunctions();
        const result = await excelApi.execute.call(executeFunctions);

        expect(result[0][0].json.success).toBe(true);
        
        const requestBody = mockFunctions.getLastRequestBody();
        expect(requestBody.lookup_column).toBe('Full Name (EN)');
      });

      it('should work with Chinese column names', async () => {
        mockFunctions.setParameter('fileName', 'chinese.xlsx');
        mockFunctions.setParameter('sheetName', '員工資料');
        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/headers?file=chinese.xlsx&sheet=%E5%93%A1%E5%B7%A5%E8%B3%87%E6%96%99',
          {
            success: true,
            headers: ['編號', '姓名', '聯絡電話', '電子郵件'],
          }
        );

        const loadFunctions = mockFunctions.getLoadOptionsFunctions();
        const columns = await excelApi.methods.loadOptions.getColumnNames.call(
          loadFunctions
        );

        expect(columns).toHaveLength(4);
        expect(columns.map(c => c.value)).toEqual(['編號', '姓名', '聯絡電話', '電子郵件']);

        // Use Chinese column name
        mockFunctions.setParameter('operation', 'update', 0);
        mockFunctions.setParameter('identifyBy', 'lookup', 0);
        mockFunctions.setParameter('lookupColumn', '電子郵件', 0);
        mockFunctions.setParameter('lookupValue', 'test@example.com', 0);
        mockFunctions.setParameter('processMode', 'first', 0);
        mockFunctions.setParameter('valuesToSet', JSON.stringify({
          '聯絡電話': '0912-345-678',
        }), 0);
        mockFunctions.setInputData([{ json: {} }]);

        mockFunctions.setRequestResponse(
          'http://localhost:8000/api/excel/update_advanced',
          {
            success: true,
            message: 'Row updated successfully',
            rows_affected: 1,
          }
        );

        const executeFunctions = mockFunctions.getExecuteFunctions();
        const result = await excelApi.execute.call(executeFunctions);

        expect(result[0][0].json.success).toBe(true);
        
        const requestBody = mockFunctions.getLastRequestBody();
        expect(requestBody.lookup_column).toBe('電子郵件');
        expect(requestBody.lookup_value).toBe('test@example.com');
      });
    });
  });
});
