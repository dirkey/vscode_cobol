import * as fs from "fs";
import { ICOBOLSettings } from "./iconfiguration";

export class COBOLFileUtils {
    static readonly isWin32 = process.platform === "win32";

    public static isFile(path: string): boolean {
        try {
            return fs.existsSync(path);
        } catch {
            return false;
        }
    }

    public static isDirectory(path: string): boolean {
        try {
            const stat = fs.statSync(path, { bigint: true });
            return stat.isDirectory();
        } catch {
            return false;
        }
    }

    public static isValidCopybookExtension(filename: string, settings: ICOBOLSettings): boolean {
        const extension = filename.includes(".") ? filename.split(".").pop()! : filename;
        return settings.copybookexts.includes(extension);
    }

    public static isValidProgramExtension(filename: string, settings: ICOBOLSettings): boolean {
        const extension = filename.includes(".") ? filename.split(".").pop()! : "";
        return settings.program_extensions.includes(extension);
    }

    public static isDirectPath(dir: string | undefined | null): boolean {
        if (!dir) return false;

        if (COBOLFileUtils.isWin32) {
            return (dir.length > 2 && dir[1] === ":") || dir.startsWith("\\");
        }

        return dir.startsWith("/");
    }

    public static isNetworkPath(dir: string | undefined | null): boolean {
        if (!dir) return false;
        return COBOLFileUtils.isWin32 && dir.startsWith("\\");
    }

    public static cleanupFilename(filename: string): string {
        let trimmed = filename.trim();
        if ((trimmed.startsWith("\"") && trimmed.endsWith("\"")) || 
            (trimmed.startsWith("'") && trimmed.endsWith("'"))) {
            return trimmed.slice(1, -1);
        }
        return trimmed;
    }
}
