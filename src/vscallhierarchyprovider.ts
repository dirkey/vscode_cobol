import * as vscode from 'vscode';
import { COBOLToken, COBOLTokenStyle, SharedSourceReferences, SourceReference_Via_Length } from './cobolsourcescanner';
import { VSCOBOLConfiguration } from './vsconfiguration';
import { VSExternalFeatures } from './vsexternalfeatures';
import { VSCOBOLSourceScanner } from './vscobolscanner';
import { ICOBOLSourceScanner } from './icobolsourcescanner';

export class COBOLHierarchyProvider implements vscode.CallHierarchyProvider {
    private current: ICOBOLSourceScanner | undefined;

    private getTargetToken(wordLower: string): COBOLToken | undefined {
        if (this.current === undefined) {
            return undefined;
        }
        
        let targetToken = this.current.sections.get(wordLower);
        if (targetToken === undefined) {
            targetToken = this.current.paragraphs.get(wordLower);
        }
        return targetToken;
    }

    private createCallHierarchyItem(name: string, detail: string, uri: vscode.Uri, range: vscode.Range): vscode.CallHierarchyItem {
        return new vscode.CallHierarchyItem(vscode.SymbolKind.Method, name, detail, uri, range, range);
    }

    prepareCallHierarchy(document: vscode.TextDocument, position: vscode.Position, token: vscode.CancellationToken): vscode.ProviderResult<vscode.CallHierarchyItem | vscode.CallHierarchyItem[]> {
        const range = document.getWordRangeAtPosition(position);
        if (!range) {
            return undefined;
        }
        
        const word = document.getText(range);
        const settings = VSCOBOLConfiguration.get_resource_settings(document, VSExternalFeatures);
        
        if (settings.enable_program_information === false) {
            return undefined;
        }
        
        this.current = VSCOBOLSourceScanner.getCachedObject(document, settings);
        if (this.current === undefined) {
            return undefined;
        }

        const sourceRefs: SharedSourceReferences = this.current.sourceReferences;
        let detail = "";
        let finalWord = word;
        let finalRange = range;

        if (!sourceRefs.targetReferences.has(word.toLowerCase())) {
            const foundToken = this.current.findNearestSectionOrParagraph(position.line);
            if (foundToken !== undefined) {
                finalWord = foundToken.tokenNameLower;
                if (foundToken.isImplicitToken) {
                    detail = foundToken.description;
                } else {
                    detail = `@${1 + foundToken.startLine}`;
                }
                finalRange = new vscode.Range(
                    new vscode.Position(foundToken.rangeStartLine, foundToken.rangeStartColumn),
                    new vscode.Position(foundToken.rangeEndLine, foundToken.rangeEndColumn)
                );
            }
        }

        return this.createCallHierarchyItem(finalWord, detail, document.uri, finalRange);
    }

    provideCallHierarchyIncomingCalls(item: vscode.CallHierarchyItem, token: vscode.CancellationToken): vscode.ProviderResult<vscode.CallHierarchyIncomingCall[]> {
        if (this.current === undefined) {
            return [];
        }

        const sourceRefs: SharedSourceReferences = this.current.sourceReferences;
        const wordLower = item.name.toLowerCase();

        if (!sourceRefs.targetReferences.has(wordLower)) {
            return undefined;
        }

        const targetRefs: SourceReference_Via_Length[] | undefined = sourceRefs.targetReferences.get(wordLower);
        const targetToken = this.getTargetToken(wordLower);
        
        if (targetToken === undefined || targetRefs === undefined) {
            return [];
        }

        const results: vscode.CallHierarchyIncomingCall[] = [];

        for (const sr of targetRefs) {
            // skip definition
            if (sr.line === targetToken.startLine && sr.column === targetToken.startColumn) {
                continue;
            }

            const uiref = vscode.Uri.parse(sourceRefs.filenameURIs[sr.fileIdentifer]);
            const range = new vscode.Range(
                new vscode.Position(sr.line, sr.column),
                new vscode.Position(sr.line, sr.column + sr.length)
            );
            const r = sr.reason.length === 0 ? sr.name : `${sr.reason} ${sr.name}`;
            const d = `@${1 + sr.line}`;
            const newitem = this.createCallHierarchyItem(r, d, uiref, range);
            const call = new vscode.CallHierarchyIncomingCall(newitem, [range]);
            results.push(call);
        }

        return results;
    }

    provideCallHierarchyOutgoingCalls(item: vscode.CallHierarchyItem, token: vscode.CancellationToken): vscode.ProviderResult<vscode.CallHierarchyOutgoingCall[]> {
        if (this.current === undefined) {
            return [];
        }

        const qp: ICOBOLSourceScanner = this.current;
        const wordLower = item.name.toLowerCase();

        if (!qp.paragraphs.has(wordLower) && !qp.sections.has(wordLower)) {
            return [];
        }

        const inToken = this.getTargetToken(wordLower);
        
        if (inToken === undefined) {
            return [];
        }

        const results: vscode.CallHierarchyOutgoingCall[] = [];
        const state = qp.sourceReferences.state;
        
        if (state.currentSectionOutRefs !== undefined && state.currentSectionOutRefs.has(inToken.tokenNameLower)) {
            const srefs = state.currentSectionOutRefs.get(inToken.tokenNameLower);
            
            if (srefs !== undefined) {
                for (const sr of srefs) {
                    if (sr.nameLower === wordLower) {
                        continue;
                    }
                    
                    if (sr.tokenStyle === COBOLTokenStyle.Section || sr.tokenStyle === COBOLTokenStyle.Paragraph) {
                        const range = new vscode.Range(
                            new vscode.Position(sr.line, sr.column),
                            new vscode.Position(sr.line, sr.column + sr.length)
                        );
                        const r = `@${1 + sr.line}`;
                        const newitem = this.createCallHierarchyItem(sr.name, r, item.uri, range);
                        const call = new vscode.CallHierarchyOutgoingCall(newitem, [range]);
                        results.push(call);
                    }
                }
            }
        }

        return results;
    }
}