import { describe, it, expect, vi, beforeEach } from 'vitest';
import path from 'path';
import { execSync } from 'child_process';
import { createRequire } from 'module';
const require = createRequire(import.meta.url);

const html2ppt = require('./index');

describe('html2ppt', () => {
  describe('isHttpUrl', () => {
    it('should return true for http/https URLs', () => {
      expect(html2ppt.isHttpUrl('http://example.com')).toBe(true);
      expect(html2ppt.isHttpUrl('https://example.com')).toBe(true);
    });
    it('should return false for others', () => {
      expect(html2ppt.isHttpUrl('file://test')).toBe(false);
      expect(html2ppt.isHttpUrl(null)).toBe(false);
    });
  });

  describe('toPageUrl', () => {
    it('should return the same URL if it is a web URL', () => {
      expect(html2ppt.toPageUrl('http://example.com')).toBe('http://example.com');
    });
    it('should return file:// URL for existing local file', () => {
      const mockFs = { existsSync: vi.fn().mockReturnValue(true) };
      const result = html2ppt.toPageUrl('package.json', mockFs);
      expect(result).toMatch(/^file:\/\//);
    });
    it('should throw error if local file does not exist', () => {
      const mockFs = { existsSync: vi.fn().mockReturnValue(false) };
      expect(() => html2ppt.toPageUrl('nonexistent.html', mockFs)).toThrow('HTML file not found');
    });
  });

  describe('renderToPng', () => {
    it('should launch browser and take screenshot of selector', async () => {
      const mockPage = {
        goto: vi.fn().mockResolvedValue(true),
        waitForTimeout: vi.fn().mockResolvedValue(true),
        locator: vi.fn().mockReturnValue({
          count: vi.fn().mockResolvedValue(1),
          first: vi.fn().mockReturnValue({
            screenshot: vi.fn().mockResolvedValue(true),
          }),
        }),
      };
      const mockBrowser = {
        newPage: vi.fn().mockResolvedValue(mockPage),
        close: vi.fn().mockResolvedValue(true),
      };
      const mockChromium = {
        launch: vi.fn().mockResolvedValue(mockBrowser),
      };

      await html2ppt.renderToPng({
        pageUrl: 'http://example.com',
        pngPath: 'out.png',
        selector: '.slide',
        width: 960,
        height: 540,
        scale: 2,
      }, mockChromium);

      expect(mockChromium.launch).toHaveBeenCalled();
      expect(mockPage.goto).toHaveBeenCalled();
      expect(mockBrowser.close).toHaveBeenCalled();
    });

    it('should take full page screenshot if selector is null', async () => {
        const mockPage = {
          goto: vi.fn().mockResolvedValue(true),
          waitForTimeout: vi.fn().mockResolvedValue(true),
          screenshot: vi.fn().mockResolvedValue(true),
        };
        const mockBrowser = {
          newPage: vi.fn().mockResolvedValue(mockPage),
          close: vi.fn().mockResolvedValue(true),
        };
        const mockChromium = {
          launch: vi.fn().mockResolvedValue(mockBrowser),
        };

        await html2ppt.renderToPng({
          pageUrl: 'http://example.com',
          pngPath: 'out.png',
          selector: null,
          width: 960,
          height: 540,
          scale: 2,
        }, mockChromium);

        expect(mockPage.screenshot).toHaveBeenCalledWith({ path: 'out.png', fullPage: true });
    });

    it('should throw error if selector is not found', async () => {
        const mockPage = {
          goto: vi.fn().mockResolvedValue(true),
          waitForTimeout: vi.fn().mockResolvedValue(true),
          locator: vi.fn().mockReturnValue({
            count: vi.fn().mockResolvedValue(0),
          }),
        };
        const mockBrowser = {
          newPage: vi.fn().mockResolvedValue(mockPage),
          close: vi.fn().mockResolvedValue(true),
        };
        const mockChromium = {
          launch: vi.fn().mockResolvedValue(mockBrowser),
        };

        await expect(html2ppt.renderToPng({
          pageUrl: 'http://example.com',
          pngPath: 'out.png',
          selector: '.notfound',
          width: 960,
          height: 540,
          scale: 2,
        }, mockChromium)).rejects.toThrow('Selector not found: .notfound');
        expect(mockBrowser.close).toHaveBeenCalled();
      });
  });

  describe('pngToPptx', () => {
    it('should create a pptx with the image', async () => {
      const mockWriteFile = vi.fn().mockResolvedValue(true);
      const mockPptxInstance = {
        addSlide: vi.fn().mockReturnValue({ addImage: vi.fn() }),
        writeFile: mockWriteFile,
        defineLayout: vi.fn(),
        layout: '',
      };
      const MockPptxGenJS = vi.fn(() => mockPptxInstance);

      await html2ppt.pngToPptx({
        pngPath: 'test.png',
        pptxPath: 'test.pptx',
        widescreen: true,
      }, MockPptxGenJS);

      expect(MockPptxGenJS).toHaveBeenCalled();
      expect(mockPptxInstance.defineLayout).toHaveBeenCalled();
      expect(mockWriteFile).toHaveBeenCalledWith({ fileName: 'test.pptx' });
    });

    it('should work without widescreen', async () => {
        const mockWriteFile = vi.fn().mockResolvedValue(true);
        const mockPptxInstance = {
          addSlide: vi.fn().mockReturnValue({ addImage: vi.fn() }),
          writeFile: mockWriteFile,
          layout: '',
        };
        const MockPptxGenJS = vi.fn(() => mockPptxInstance);

        await html2ppt.pngToPptx({
          pngPath: 'test.png',
          pptxPath: 'test.pptx',
          widescreen: false,
        }, MockPptxGenJS);

        expect(mockWriteFile).toHaveBeenCalled();
    });
  });

  describe('main', () => {
    it('should run the full process', async () => {
      const originalArgv = process.argv;
      process.argv = ['node', 'index.js', 'http://example.com', '--out', 'test.pptx', '--png', 'test.png', '--selector', '.slide'];

      const toPageUrlSpy = vi.spyOn(html2ppt, 'toPageUrl').mockReturnValue('http://example.com');
      const renderToPngSpy = vi.spyOn(html2ppt, 'renderToPng').mockResolvedValue(true);
      const pngToPptxSpy = vi.spyOn(html2ppt, 'pngToPptx').mockResolvedValue(true);
      const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});

      await html2ppt.main();

      expect(toPageUrlSpy).toHaveBeenCalledWith('http://example.com');
      expect(renderToPngSpy).toHaveBeenCalled();
      expect(pngToPptxSpy).toHaveBeenCalled();

      process.argv = originalArgv;
      toPageUrlSpy.mockRestore();
      renderToPngSpy.mockRestore();
      pngToPptxSpy.mockRestore();
      logSpy.mockRestore();
    });

    it('should handle missing input gracefully', async () => {
        const originalArgv = process.argv;
        process.argv = ['node', 'index.js'];
        const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});
        const exitSpy = vi.spyOn(process, 'exit').mockImplementation(() => {});

        await html2ppt.main();

        process.argv = originalArgv;
        logSpy.mockRestore();
        exitSpy.mockRestore();
    });

    it('should handle default png path and empty selector', async () => {
        const originalArgv = process.argv;
        process.argv = ['node', 'index.js', 'http://example.com', '--out', 'test.pptx', '--selector', ' '];

        const toPageUrlSpy = vi.spyOn(html2ppt, 'toPageUrl').mockReturnValue('http://example.com');
        const renderToPngSpy = vi.spyOn(html2ppt, 'renderToPng').mockResolvedValue(true);
        const pngToPptxSpy = vi.spyOn(html2ppt, 'pngToPptx').mockResolvedValue(true);
        const logSpy = vi.spyOn(console, 'log').mockImplementation(() => {});

        await html2ppt.main();

        expect(renderToPngSpy).toHaveBeenCalledWith(expect.objectContaining({
            selector: null
        }));

        process.argv = originalArgv;
        toPageUrlSpy.mockRestore();
        renderToPngSpy.mockRestore();
        pngToPptxSpy.mockRestore();
        logSpy.mockRestore();
    });
  });

  describe('CLI', () => {
    it('should run as a CLI', () => {
        const output = execSync('node index.js --help', { stdio: 'pipe' });
        expect(output.toString()).toContain('Usage:');
    });

    it('should be covered even if run as module', async () => {
        // This is to cover the branch if (require.main === module)
        // But we can't easily make it true without actually running it.
        // The execSync above already runs it.
    });
  });
});
