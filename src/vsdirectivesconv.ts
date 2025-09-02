export class PortResult {
  constructor(
    public readonly filename: string,
    public readonly linenum: number,
    public readonly replaceLine: string,
    public readonly message: string
  ) {}
}

class PortSwap {
  constructor(
    public lastNumber: number,
    public search: RegExp,
    public replaceLine: string,
    public message: string
  ) {}
}

// ---------------- IBM to Micro Focus Mapping ----------------
// https://www.ibm.com/docs/en/cobol-zos/6.1?topic=program-compiler-options
// https://www.microfocus.com/documentation/enterprise-developer/ed90/ED-EclipseUNIX/HRCDRHCDIR0W.html
export const ibm2mf: [string, string][] = [
  ["QUOTE","QUOTE"],
  ["APOST","APOST"],
  ["CICS","CICSECM"],
  ["CODEPAGE(XX)",""],
  ["CURRENCY",""],
  ["NOCURRENCY",""],
  ["NSYMBOL(NATIONAL)","NSYMBOL(NATIONAL)"],
  ["NS(DBDC)","NSYMBOL(DBCS)"],
  ["NS(NAT)","NSYMBOL(NATIONAL)"],
  ["NUMBER",""],
  ["NONUMBER",""],
  ["QUALIFY(COMPAT)",""],
  ["QUA(C)",""],
  ["QUA(E)",""],
  ["SEQUENCE","SQCHK"],
  ["SEQ","SQCHK"],
  ["NOSEQ","NOSQCHK"],
  ["SQL","SQL"],
  ["NOSQL","NOSQL"],
  ["SQLCCSID",""],
  ["SQLC",""],
  ["SQLIMS",""],
  ["SUPPRESS",""],
  ["SUPPR",""],
  ["WORD","ADDRSV"],
  ["NOWORD",""],
  ["XMLPARSE(XMLSS","XMLPARSE(XMLSS)"],
  ["XP(X)","XMLPARSE(XMLSS)"],
  ["XP(C)","XMLPARSE(COMPAT)"],
  ["INTDATE(ANSI)","INTDATE(ANSI)"],
  ["INTDATE(LILIAN)","INTDATE(LILIAN)"],
  ["LANGUAGE(ENGLISH)",""],
  ["LANG(EN)","NOJAPANESE"],
  ["LANG(UE)",""],
  ["LANG(JA)","JAPANESE"],
  ["LANG(JP)","JAPANESE"],
  ["LINECOUNT(60)","LINE-COUNT(60)"],
  ["LC(60)","LINE-COUNT(60)"],
  ["LIST","LIST"],
  ["NOLIST","NOLIST"],
  ["MAP","XREF"],
  ["NOMAP","NOXREF"],
];

// ---------------- SourcePorter ----------------
interface DirectiveConfig {
  lastNumber: number;
  search: RegExp;
  replaceLine: string;
  message: string;
}

export class SourcePorter {
  private readonly portsDirectives: PortSwap[];

  // Fully declarative directive definitions
  private static readonly DIRECTIVES: DirectiveConfig[] = [
    { lastNumber: 100, search: />>CALL-CONVENTION COBOL/i, replaceLine: "$set defaultcalls(0)", message: "Change to defaultcalls(0)" },
    { lastNumber: 100, search: /(>>CALL-CONVENTION EXTERN)/i, replaceLine: "*> $1", message: "Comment out" },
    { lastNumber: 100, search: />>CALL-CONVENTION STDCALL/i, replaceLine: "$set defaultcalls(74)", message: "Change to defaultcalls(74)" },
    { lastNumber: 100, search: />>CALL-CONVENTION STATIC/i, replaceLine: "$set litlink", message: "Change to $set litlink" },
    { lastNumber: 100, search: />>SOURCE\s+FORMAT\s+(IS\s+FREE|FREE)/i, replaceLine: "$set sourceformat(free)", message: "Change to $set sourceformat(free)" },
    { lastNumber: 100, search: />>SOURCE\s+FORMAT\s+(IS\s+FIXED|FIXED)/i, replaceLine: "$set sourceformat(fixed)", message: "Change to $set sourceformat(fixed)" },
    { lastNumber: 100, search: />>SOURCE\s+FORMAT\s+(IS\s+VARIABLE|VARIABLE)/i, replaceLine: "$set sourceformat(variable)", message: "Change to $set sourceformat(variable)" },
    { lastNumber: Number.MAX_VALUE, search: /FUNCTION\s+SUBSTITUTE/i, replaceLine: "", message: "Re-write to use 'INSPECT REPLACING' or custom/search replace" },
  ];

  constructor() {
    // Map declarative directives to PortSwap instances
    this.portsDirectives = SourcePorter.DIRECTIVES.map(
      ({ lastNumber, search, replaceLine, message }) => new PortSwap(lastNumber, search, replaceLine, message)
    );
  }

  public isDirectiveChangeRequired(filename: string, lineNumber: number, line: string): PortResult | undefined {
    // First five lines are always active
    if (lineNumber <= 5) return undefined;

    const directive = this.portsDirectives.find(d => d.lastNumber >= lineNumber && d.search.test(line));
    if (!directive) return undefined;

    const replaced = directive.replaceLine ? line.replace(directive.search, directive.replaceLine) : "";
    return new PortResult(filename, lineNumber, replaced, directive.message);
  }
}