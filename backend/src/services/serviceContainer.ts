import { ExcelParserService } from "./excelParserService";
import { FormulaParserService } from "./formulaParserService";
import { GraphBuilderService } from "./graphBuilderService";
import { ValidationService } from "./validationService";
import { ExecutionEngineService } from "./executionEngineService";
import { WorkbookSessionService } from "./workbookSessionService";
import { ExportService } from "./exportService";
import { FileRegistryService } from "./fileRegistryService";

export const formulaParserService = new FormulaParserService();
export const graphBuilderService = new GraphBuilderService(formulaParserService);
export const excelParserService = new ExcelParserService();
export const validationService = new ValidationService();
export const executionEngineService = new ExecutionEngineService();
export const workbookSessionService = new WorkbookSessionService();
export const exportService = new ExportService();
export const fileRegistryService = new FileRegistryService();
