import { ExcelApiAuth } from '../../credentials/ExcelApiAuth.credentials';

describe('ExcelApiAuth Credentials', () => {
  let credentials: ExcelApiAuth;

  beforeEach(() => {
    credentials = new ExcelApiAuth();
  });

  it('should have correct credential name', () => {
    expect(credentials.name).toBe('excelApiAuth');
  });

  it('should have correct display name', () => {
    expect(credentials.displayName).toBe('Excel API Auth');
  });

  it('should have documentation URL', () => {
    expect(credentials.documentationUrl).toBeDefined();
    expect(typeof credentials.documentationUrl).toBe('string');
  });

  it('should have required properties', () => {
    expect(credentials.properties).toHaveLength(2);
  });

  describe('API URL Property', () => {
    it('should have correct configuration', () => {
      const urlProperty = credentials.properties.find(p => p.name === 'url');
      
      expect(urlProperty).toBeDefined();
      expect(urlProperty?.displayName).toBe('API URL');
      expect(urlProperty?.type).toBe('string');
      expect(urlProperty?.required).toBe(true);
      expect(urlProperty?.default).toBe('http://localhost:8000');
    });
  });

  describe('API Token Property', () => {
    it('should have correct configuration', () => {
      const tokenProperty = credentials.properties.find(p => p.name === 'token');
      
      expect(tokenProperty).toBeDefined();
      expect(tokenProperty?.displayName).toBe('API Token');
      expect(tokenProperty?.type).toBe('string');
      expect(tokenProperty?.required).toBe(true);
      expect(tokenProperty?.typeOptions?.password).toBe(true);
    });
  });
});
