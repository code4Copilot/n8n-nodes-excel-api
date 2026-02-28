import { ExcelApi } from '../../../nodes/ExcelApi/ExcelApi.node';
import { MockExecuteFunctions } from '../../helpers/mockHelpers';

describe('ExcelApi Node - Type Conversion', () => {
  let excelApi: ExcelApi;
  let mockFunctions: MockExecuteFunctions;

  beforeEach(() => {
    excelApi = new ExcelApi();
    mockFunctions = new MockExecuteFunctions();
    
    mockFunctions.setCredentials('excelApiAuth', {
      url: 'http://localhost:8000',
      token: 'test-token-123',
    });

    mockFunctions.setParameter('operation', 'append', 0);
    mockFunctions.setParameter('fileName', 'test.xlsx', 0);
    mockFunctions.setParameter('sheetName', 'Sheet1', 0);
    mockFunctions.setParameter('appendMode', 'object', 0);
    mockFunctions.setInputData([{ json: {} }]);
  });

  describe('String to Number Conversion', () => {
    it('should convert integer string to number', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Quantity': '123',
        'Name': 'Test',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Quantity']).toBe(123);
      expect(typeof capturedBody.values['Quantity']).toBe('number');
      expect(capturedBody.values['Name']).toBe('Test');
    });

    it('should convert float string to number', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Price': '45.67',
        'Discount': '0.85',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Price']).toBe(45.67);
      expect(typeof capturedBody.values['Price']).toBe('number');
      expect(capturedBody.values['Discount']).toBe(0.85);
    });

    it('should convert negative numbers', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Temperature': '-15.5',
        'Balance': '-100',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Temperature']).toBe(-15.5);
      expect(capturedBody.values['Balance']).toBe(-100);
    });

    it('should not convert invalid number strings', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Code': '123abc',
        'ID': 'A-456',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Code']).toBe('123abc');
      expect(capturedBody.values['ID']).toBe('A-456');
    });
  });

  describe('String to Boolean Conversion', () => {
    it('should convert "true" string to boolean', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Active': 'true',
        'Verified': 'True',
        'Completed': 'TRUE',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Active']).toBe(true);
      expect(capturedBody.values['Verified']).toBe(true);
      expect(capturedBody.values['Completed']).toBe(true);
    });

    it('should convert "false" string to boolean', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Disabled': 'false',
        'Deleted': 'False',
        'Closed': 'FALSE',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Disabled']).toBe(false);
      expect(capturedBody.values['Deleted']).toBe(false);
      expect(capturedBody.values['Closed']).toBe(false);
    });
  });

  describe('String to Null Conversion', () => {
    it('should convert "null" string to null', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Note': 'null',
        'Comment': 'Null',
        'Description': 'NULL',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Note']).toBe(null);
      expect(capturedBody.values['Comment']).toBe(null);
      expect(capturedBody.values['Description']).toBe(null);
    });

    it('should convert empty string to null', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'OptionalField1': '',
        'OptionalField2': '  ',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['OptionalField1']).toBe(null);
      expect(capturedBody.values['OptionalField2']).toBe(null);
    });

    it('should keep null values as null', async () => {
      const values = {
        'Field1': null,
        'Field2': undefined,
      };
      mockFunctions.setParameter('appendValuesObject', values, 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Field1']).toBe(null);
      expect(capturedBody.values['Field2']).toBe(null);
    });
  });

  describe('Date Conversion', () => {
    it('should preserve date-only string without adding time', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'CreatedDate': '2024-01-15',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['CreatedDate']).toBe('2024-01-15');
    });

    it('should convert ISO datetime string', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'UpdatedAt': '2024-01-15T10:30:00',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['UpdatedAt']).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
    });

    it('should convert ISO datetime with timezone', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'Timestamp': '2024-01-15T10:30:00.000Z',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Timestamp']).toMatch(/Z$/);
    });
  });

  describe('Array Mode Type Conversion', () => {
    it('should convert types in array mode', async () => {
      mockFunctions.setParameter('appendMode', 'array', 0);
      mockFunctions.setParameter('appendValuesArray', JSON.stringify([
        'E001',
        '123',
        'true',
        '45.67',
        'null',
        '2024-01-15',
      ]), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values[0]).toBe('E001');
      expect(capturedBody.values[1]).toBe(123);
      expect(capturedBody.values[2]).toBe(true);
      expect(capturedBody.values[3]).toBe(45.67);
      expect(capturedBody.values[4]).toBe(null);
      expect(capturedBody.values[5]).toBe('2024-01-15');
    });
  });

  describe('Update Operation Type Conversion', () => {
    beforeEach(() => {
      mockFunctions.setParameter('operation', 'update', 0);
      mockFunctions.setParameter('identifyBy', 'rowNumber', 0);
      mockFunctions.setParameter('rowNumber', 2, 0);
    });

    it('should convert types in update operation', async () => {
      mockFunctions.setParameter('valuesToSet', JSON.stringify({
        'Quantity': '500',
        'Active': 'true',
        'Price': '99.99',
        'Note': 'null',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/update_advanced',
        {
          success: true,
          message: 'Row updated successfully',
          updated_count: 1,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values_to_set['Quantity']).toBe(500);
      expect(capturedBody.values_to_set['Active']).toBe(true);
      expect(capturedBody.values_to_set['Price']).toBe(99.99);
      expect(capturedBody.values_to_set['Note']).toBe(null);
    });
  });

  describe('Mixed Data Types', () => {
    it('should handle mixed data types correctly', async () => {
      mockFunctions.setParameter('appendValuesObject', JSON.stringify({
        'EmployeeID': 'E001',
        'Name': 'John Doe',
        'Age': '30',
        'Salary': '50000.50',
        'IsActive': 'true',
        'TerminationDate': 'null',
        'HireDate': '2020-01-15',
        'Department': 'IT',
      }), 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['EmployeeID']).toBe('E001');
      expect(capturedBody.values['Name']).toBe('John Doe');
      expect(capturedBody.values['Age']).toBe(30);
      expect(capturedBody.values['Salary']).toBe(50000.50);
      expect(capturedBody.values['IsActive']).toBe(true);
      expect(capturedBody.values['TerminationDate']).toBe(null);
      expect(capturedBody.values['HireDate']).toBe('2020-01-15');
      expect(capturedBody.values['Department']).toBe('IT');
    });
  });

  describe('Already Typed Values', () => {
    it('should preserve already typed values', async () => {
      const values = {
        'Number': 100,
        'Boolean': true,
        'String': 'Keep as string',
        'Null': null,
      };
      mockFunctions.setParameter('appendValuesObject', values, 0);

      mockFunctions.setRequestResponse(
        'http://localhost:8000/api/excel/append_object',
        {
          success: true,
          message: 'Row appended successfully',
          row: 2,
        }
      );

      const executeFunctions = mockFunctions.getExecuteFunctions();
      await excelApi.execute.call(executeFunctions);

      const capturedBody = mockFunctions.getLastRequestBody();
      expect(capturedBody.values['Number']).toBe(100);
      expect(capturedBody.values['Boolean']).toBe(true);
      expect(capturedBody.values['String']).toBe('Keep as string');
      expect(capturedBody.values['Null']).toBe(null);
    });
  });
});
