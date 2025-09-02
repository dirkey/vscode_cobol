/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-module-boundary-types */

import path from "path";
import { COBOLFileSymbol, COBOLWorkspaceFile } from "./cobolglobalcache";
import { IExternalFeatures } from "./externalfeatures";
import { InMemoryGlobalCacheHelper, InMemoryGlobalSymbolCache } from "./globalcachehelper";
import { ICOBOLSettings } from "./iconfiguration";
import { COBOLFileUtils } from "./fileutils";

export enum TypeCategory {
    ClassId = "T",
    InterfaceId = "I",
    EnumId = "E"
}

export class COBOLWorkspaceSymbolCacheHelper {

    private static literalRegex = /^([a-zA-Z0-9_-]*[a-zA-Z0-9]|([#]?[0-9a-zA-Z]+[a-zA-Z0-9_-]*[a-zA-Z0-9]))$/;

    private static isValidLiteral(id: string): boolean {
        return !!id && COBOLWorkspaceSymbolCacheHelper.literalRegex.test(id);
    }

    private static removeAllProgramSymbols(srcfilename: string, symbolsCache: Map<string, COBOLFileSymbol[]>) {
        for (const [key, symbolList] of symbolsCache) {
            if (!symbolList) continue;
            const newSymbols = symbolList.filter(s => s.filename !== srcfilename);
            if (newSymbols.length) {
                symbolsCache.set(key, newSymbols);
            } else {
                symbolsCache.delete(key);
            }
        }
    }

    private static addSymbolToCache(
        srcfilename: string,
        symbolUnchanged: string,
        lineNumber: number,
        symbolsCache: Map<string, COBOLFileSymbol[]>
    ) {
        const symbol = symbolUnchanged.toLowerCase();
        const list = symbolsCache.get(symbol) || [];
        const existingSymbols = list.filter(s => s.filename === srcfilename);

        if (existingSymbols.length === 0) {
            list.push(new COBOLFileSymbol(srcfilename, lineNumber));
        } else if (existingSymbols.length === 1) {
            existingSymbols[0].linenum = lineNumber;
        } else {
            const nonFileSymbol = existingSymbols.find(s => s.linenum !== 1);
            if (nonFileSymbol) nonFileSymbol.linenum = lineNumber;
        }

        symbolsCache.set(symbol, list);
        InMemoryGlobalSymbolCache.isDirty = true;
    }

    public static addCalableSymbol(srcfilename: string, symbolUnchanged: string, lineNumber: number) {
        if (!srcfilename || !symbolUnchanged) return;

        const fileName = COBOLFileUtils.cleanupFilename(InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename));
        const fileNameNoExt = path.basename(fileName, path.extname(fileName));
        const callableSymbolFromFilenameLower = fileNameNoExt.toLowerCase();

        if (!COBOLWorkspaceSymbolCacheHelper.isValidLiteral(symbolUnchanged)) return;

        if (symbolUnchanged.toLowerCase() === callableSymbolFromFilenameLower) {
            InMemoryGlobalSymbolCache.defaultCallableSymbols.set(callableSymbolFromFilenameLower, srcfilename);
            return;
        }

        if (InMemoryGlobalSymbolCache.defaultCallableSymbols.has(callableSymbolFromFilenameLower)) {
            InMemoryGlobalSymbolCache.defaultCallableSymbols.delete(callableSymbolFromFilenameLower);
        }

        COBOLWorkspaceSymbolCacheHelper.addSymbolToCache(fileName, symbolUnchanged, lineNumber, InMemoryGlobalSymbolCache.callableSymbols);
    }

    public static addEntryPoint(srcfilename: string, symbolUnchanged: string, lineNumber: number) {
        if (!srcfilename || !symbolUnchanged) return;
        const filename = COBOLFileUtils.cleanupFilename(InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename));
        COBOLWorkspaceSymbolCacheHelper.addSymbolToCache(filename, symbolUnchanged, lineNumber, InMemoryGlobalSymbolCache.entryPoints);
    }

    public static addReferencedCopybook(copybook: string, fullInFilename: string) {
        const inFilename = COBOLFileUtils.cleanupFilename(InMemoryGlobalCacheHelper.getFilenameWithoutPath(fullInFilename));
        const encodedKey = `${copybook},${inFilename}`;
        if (!InMemoryGlobalSymbolCache.knownCopybooks.has(encodedKey)) {
            InMemoryGlobalSymbolCache.knownCopybooks.set(encodedKey, copybook);
            InMemoryGlobalSymbolCache.isDirty = true;
        }
    }

    public static removeAllCopybookReferences(fullInFilename: string) {
        const inFilename = InMemoryGlobalCacheHelper.getFilenameWithoutPath(fullInFilename);
        for (const key of [...InMemoryGlobalSymbolCache.knownCopybooks.keys()]) {
            if (key.split(",")[1] === inFilename) InMemoryGlobalSymbolCache.knownCopybooks.delete(key);
        }
    }

    public static addClass(srcfilename: string, symbolUnchanged: string, lineNumber: number, category: TypeCategory) {
        const map = category === TypeCategory.InterfaceId
            ? InMemoryGlobalSymbolCache.interfaces
            : category === TypeCategory.EnumId
                ? InMemoryGlobalSymbolCache.enums
                : InMemoryGlobalSymbolCache.types;

        COBOLWorkspaceSymbolCacheHelper.addSymbolToCache(
            InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename),
            symbolUnchanged,
            lineNumber,
            map
        );
    }

    public static removeAllPrograms(srcfilename: string) {
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramSymbols(
            InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename),
            InMemoryGlobalSymbolCache.callableSymbols
        );
    }

    public static removeAllProgramEntryPoints(srcfilename: string) {
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramSymbols(
            InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename),
            InMemoryGlobalSymbolCache.entryPoints
        );
    }

    public static removeAllTypes(srcfilename: string) {
        const f = InMemoryGlobalCacheHelper.getFilenameWithoutPath(srcfilename);
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramSymbols(f, InMemoryGlobalSymbolCache.types);
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramSymbols(f, InMemoryGlobalSymbolCache.enums);
        COBOLWorkspaceSymbolCacheHelper.removeAllProgramSymbols(f, InMemoryGlobalSymbolCache.interfaces);
    }

    public static loadGlobalCacheFromArray(settings: ICOBOLSettings, symbols: string[], clear: boolean) {
        if (clear) InMemoryGlobalSymbolCache.callableSymbols.clear();

        for (const symbol of symbols) {
            const parts = symbol.split(",");
            if (parts.length === 2) COBOLWorkspaceSymbolCacheHelper.addCalableSymbol(parts[1], parts[0], 0);
            else if (parts.length === 3) COBOLWorkspaceSymbolCacheHelper.addCalableSymbol(parts[1], parts[0], Number(parts[2]));
        }
    }

    public static loadGlobalEntryCacheFromArray(settings: ICOBOLSettings, symbols: string[], clear: boolean) {
        if (clear) InMemoryGlobalSymbolCache.entryPoints.clear();
        for (const symbol of symbols) {
            const parts = symbol.split(",");
            if (parts.length === 3) COBOLWorkspaceSymbolCacheHelper.addEntryPoint(parts[1], parts[0], Number(parts[2]));
        }
    }

    public static loadGlobalKnownCopybooksFromArray(settings: ICOBOLSettings, copybookValues: string[], clear: boolean) {
        if (clear) InMemoryGlobalSymbolCache.knownCopybooks.clear();
        for (const symbol of copybookValues) {
            const parts = symbol.split(",");
            if (parts.length === 2) InMemoryGlobalSymbolCache.knownCopybooks.set(symbol, parts[1]);
        }
    }

    public static loadGlobalTypesCacheFromArray(settings: ICOBOLSettings, symbols: string[], clear: boolean) {
        if (clear) {
            InMemoryGlobalSymbolCache.enums.clear();
            InMemoryGlobalSymbolCache.interfaces.clear();
            InMemoryGlobalSymbolCache.types.clear();
        }

        for (const symbol of symbols) {
            const parts = symbol.split(",");
            if (parts.length === 4) {
                let category: TypeCategory = TypeCategory.ClassId;
                switch (parts[0]) {
                    case "I": category = TypeCategory.InterfaceId; break;
                    case "T": category = TypeCategory.ClassId; break;
                    case "E": category = TypeCategory.EnumId; break;
                }
                COBOLWorkspaceSymbolCacheHelper.addClass(parts[2], parts[1], Number(parts[3]), category);
            }
        }
    }

    public static loadFileCacheFromArray(settings: ICOBOLSettings, externalFeatures: IExternalFeatures, files: string[], clear: boolean) {
        if (clear) InMemoryGlobalSymbolCache.sourceFilenameModified.clear();

        for (const symbol of files) {
            const parts = symbol.split(",");
            if (parts.length === 2) {
                const ms = BigInt(parts[0]);
                const fullDir = externalFeatures.getFullWorkspaceFilename(parts[1], ms, settings);
                if (fullDir) {
                    InMemoryGlobalSymbolCache.sourceFilenameModified.set(fullDir, new COBOLWorkspaceFile(ms, parts[1]));
                } else {
                    COBOLWorkspaceSymbolCacheHelper.removeAllProgramEntryPoints(parts[1]);
                    COBOLWorkspaceSymbolCacheHelper.removeAllTypes(parts[1]);
                }
            }
        }
    }
}