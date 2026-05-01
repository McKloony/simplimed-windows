<#
.SYNOPSIS
    Patches globally-installed gitnexus 1.6.3 for VB6 support + Windows compatibility.
.DESCRIPTION
    Idempotent -- safe to run multiple times or after 'npm install -g gitnexus'.
    Run AFTER installing gitnexus:
        npm install -g gitnexus
        powershell -ExecutionPolicy Bypass -File .\gitnexus-vb6-patch.ps1
#>

$ErrorActionPreference = 'Stop'
$enc = [System.Text.UTF8Encoding]::new($false)

function ok($msg)   { Write-Host "  [ok]   $msg" -ForegroundColor Green }
function skip($msg) { Write-Host "  [skip] $msg" -ForegroundColor Gray }
function hdr($msg)  { Write-Host "`n$msg" -ForegroundColor Cyan }

# Apply a text replacement.
# $done = string that ONLY exists AFTER patching (skip if found).
# $old  = text to replace (if not found: already patched or wrong version -> skip).
function Patch {
    param([string]$file, [string]$done, [string]$old, [string]$new, [string]$desc)
    $c = [IO.File]::ReadAllText($file, $enc)
    if ($done -and $c.Contains($done))  { skip $desc; return }
    if (-not $c.Contains($old))         { skip $desc; return }
    [IO.File]::WriteAllText($file, $c.Replace($old, $new), $enc)
    ok $desc
}

# ---------- locate gitnexus -------------------------------------------------
hdr "Locating gitnexus..."
$root = & npm root -g 2>$null
if (-not $root -or -not (Test-Path "$root\gitnexus")) {
    Write-Error "gitnexus not found. Run: npm install -g gitnexus"; exit 1
}
$dist = "$root\gitnexus\dist"
ok "Found: $dist"
$ver = (Get-Content "$root\gitnexus\package.json" -Raw | ConvertFrom-Json).version
if ($ver -ne '1.6.3') {
    Write-Host "  [warn] Expected 1.6.3, found $ver -- patches may not apply cleanly" -ForegroundColor Yellow
}

# ---------- patches ---------------------------------------------------------
hdr "Patching..."

# 1. languages.js -- add VisualBasic6 enum entry
$f = "$dist\_shared\languages.js"
$o = 'SupportedLanguages["Cobol"] = "cobol";' + [char]10 + '})(SupportedLanguages'
$n = 'SupportedLanguages["Cobol"] = "cobol";' + [char]10 `
   + '    /** Standalone regex processor for VB6 -- .bas, .frm, .cls files. */' + [char]10 `
   + '    SupportedLanguages["VisualBasic6"] = "vb6";' + [char]10 `
   + '})(SupportedLanguages'
Patch $f 'VisualBasic6' $o $n "languages.js -- add VisualBasic6 enum"

# 2. pipeline-phases/index.js -- export vb6Phase
$f = "$dist\core\ingestion\pipeline-phases\index.js"
$o = "export { cobolPhase } from './cobol.js';"
$n = "export { cobolPhase } from './cobol.js';" + [char]10 + "export { vb6Phase } from './vb6.js';"
Patch $f 'vb6Phase' $o $n "pipeline-phases/index.js -- export vb6Phase"

# 3a. pipeline.js -- add vb6Phase to import line
$f = "$dist\core\ingestion\pipeline.js"
Patch $f 'cobolPhase, vb6Phase, parsePhase,' `
    'cobolPhase, parsePhase,' `
    'cobolPhase, vb6Phase, parsePhase,' `
    "pipeline.js -- add vb6Phase to import"

# 3b. pipeline.js -- add vb6Phase to buildPhaseList array
$o3b = "        cobolPhase," + [char]10 + "        parsePhase,"
$n3b = "        cobolPhase," + [char]10 + "        vb6Phase," + [char]10 + "        parsePhase,"
$d3b = "        vb6Phase," + [char]10 + "        parsePhase,"
Patch $f $d3b $o3b $n3b "pipeline.js -- add vb6Phase to phase list"

# 4. lbug-adapter.js -- remove loadVectorExtension() call (crashes on Windows)
$f = "$dist\core\lbug\lbug-adapter.js"
$o4 = "    // Load VECTOR extension for semantic search support" + [char]10 `
    + "    await loadVectorExtension();" + [char]10 `
    + "    currentDbPath"
$n4 = "    // Load VECTOR extension for semantic search support" + [char]10 `
    + [char]10 `
    + "    currentDbPath"
Patch $f '' $o4 $n4 "lbug-adapter.js -- disable VECTOR (crashes on Windows)"

# 5. pool-adapter.js -- replace FTS loading with true (crashes on Windows)
$f = "$dist\core\lbug\pool-adapter.js"
$o5 = "        shared.ftsLoaded = await loadFTSExtension(available[0]);"
$n5 = "        shared.ftsLoaded = true; // FTS disabled -- native DLL crashes on this Windows setup"
Patch $f 'FTS disabled' $o5 $n5 "pool-adapter.js -- disable FTS (crashes on Windows)"

# ---------- new files -------------------------------------------------------
hdr "Writing new files..."

# 6. vb6-processor.js
$f = "$dist\core\ingestion\vb6-processor.js"
if (Test-Path $f) { skip "vb6-processor.js (already exists)" } else {
$c6 = @'
/**
 * VB6 Processor -- standalone regex-based, no tree-sitter.
 * Handles .bas (modules), .frm (forms), .cls (classes).
 * Schema: Module(id,name,filePath,startLine,endLine,content,description)
 *         Function(id,name,filePath,startLine,endLine,isExported,content,description)
 */
import path from 'node:path';
import { generateId } from '../../lib/utils.js';

const VB6_EXTENSIONS = new Set(['.bas', '.frm', '.cls']);

export function isVb6File(filePath) {
    return VB6_EXTENSIONS.has(path.extname(filePath).toLowerCase());
}

const PROC_DECL_RE = /^[ \t]*(?:(?:Public|Private|Friend|Static)[ \t]+)*(?:Static[ \t]+)?(Function|Sub|Property[ \t]+(?:Get|Let|Set))[ \t]+(\w+)[ \t]*[(\r\n]/gim;
const PROC_END_RE  = /^[ \t]*End[ \t]+(?:Function|Sub|Property)/gim;

function extractProcedures(content) {
    const procs = [];
    const declRe = new RegExp(PROC_DECL_RE.source, 'gim');
    let m;
    while ((m = declRe.exec(content)) !== null) {
        const startLine = content.slice(0, m.index).split('\n').length;
        const kind = m[1].replace(/\s+/g, ' ').trim();
        const name = m[2];
        procs.push({ name, kind, startLine, endLine: startLine });
    }
    const endRe = new RegExp(PROC_END_RE.source, 'gim');
    const ends = [];
    while ((m = endRe.exec(content)) !== null) {
        ends.push(content.slice(0, m.index).split('\n').length);
    }
    procs.forEach((proc) => {
        const nextEnd = ends.find(e => e > proc.startLine);
        proc.endLine = nextEnd || proc.startLine + 5;
    });
    return procs;
}

function isPublicProc(content, procName) {
    const re = new RegExp(`^[ \\t]*(?:Private|Friend)[ \\t]+(?:Function|Sub|Property)[ \\t]+${procName}\\b`, 'im');
    return !re.test(content);
}

export function processVb6(graph, files) {
    const result = { modules: 0, procedures: 0, calls: 0 };

    const allProcs = new Map();
    const fileProcs = new Map();
    for (const file of files) {
        const procs = extractProcedures(file.content);
        fileProcs.set(file.path, procs);
        for (const p of procs) {
            allProcs.set(p.name.toLowerCase(), { filePath: file.path, name: p.name });
        }
    }
    const knownNames = new Set(allProcs.keys());

    for (const file of files) {
        const fileNodeId = generateId('File', file.path);
        if (!graph.getNode(fileNodeId)) continue;

        const moduleName = path.basename(file.path, path.extname(file.path));
        const moduleId   = generateId('Module', file.path);
        const lines      = file.content.split('\n');

        graph.addNode({
            id: moduleId,
            label: 'Module',
            properties: {
                name:        moduleName,
                filePath:    file.path,
                startLine:   1,
                endLine:     lines.length,
                content:     '',
                description: '',
            },
        });
        graph.addRelationship({
            id:       generateId('CONTAINS', `${fileNodeId}->${moduleId}`),
            type:     'CONTAINS',
            sourceId: fileNodeId,
            targetId: moduleId,
            properties: {},
        });
        result.modules++;

        const procs = fileProcs.get(file.path) || [];
        for (const proc of procs) {
            const procId    = generateId('Function', `${file.path}:${proc.name}`);
            const bodyLines = lines.slice(proc.startLine, proc.endLine);
            graph.addNode({
                id: procId,
                label: 'Function',
                properties: {
                    name:        proc.name,
                    filePath:    file.path,
                    startLine:   proc.startLine,
                    endLine:     proc.endLine,
                    isExported:  isPublicProc(file.content, proc.name),
                    content:     bodyLines.slice(0, 20).join('\n'),
                    description: `${proc.kind} ${proc.name}`,
                },
            });
            graph.addRelationship({
                id:       generateId('CONTAINS', `${moduleId}->${procId}`),
                type:     'CONTAINS',
                sourceId: moduleId,
                targetId: procId,
                properties: {},
            });
            result.procedures++;
        }

        for (const proc of procs) {
            const procId    = generateId('Function', `${file.path}:${proc.name}`);
            const bodyLines = lines.slice(proc.startLine, proc.endLine).join('\n');
            const callRe    = /\b(\w+)\s*\(/g;
            let cm;
            const seen = new Set();
            while ((cm = callRe.exec(bodyLines)) !== null) {
                const calledLower = cm[1].toLowerCase();
                if (!knownNames.has(calledLower) || seen.has(calledLower)) continue;
                if (calledLower === proc.name.toLowerCase()) continue;
                seen.add(calledLower);
                const target   = allProcs.get(calledLower);
                const targetId = generateId('Function', `${target.filePath}:${target.name}`);
                graph.addRelationship({
                    id:       generateId('CALLS', `${procId}->${targetId}`),
                    type:     'CALLS',
                    sourceId: procId,
                    targetId: targetId,
                    properties: {},
                });
                result.calls++;
            }
        }
    }

    return result;
}
'@
[IO.File]::WriteAllText($f, $c6, $enc)
ok "vb6-processor.js"
}

# 7. pipeline-phases/vb6.js
$f = "$dist\core\ingestion\pipeline-phases\vb6.js"
if (Test-Path $f) { skip "pipeline-phases/vb6.js (already exists)" } else {
$c7 = @'
/**
 * Phase: vb6 -- processes VB6 files (.bas, .frm, .cls) via regex (no tree-sitter).
 */
import { getPhaseOutput } from './types.js';
import { processVb6, isVb6File } from '../vb6-processor.js';
import { readFileContents } from '../filesystem-walker.js';
import { isDev } from '../utils/env.js';

export const vb6Phase = {
    name: 'vb6',
    deps: ['structure'],
    async execute(ctx, deps) {
        const { scannedFiles } = getPhaseOutput(deps, 'structure');
        const vb6Scanned = scannedFiles.filter((f) => isVb6File(f.path));
        if (vb6Scanned.length === 0) {
            return { modules: 0, procedures: 0, calls: 0 };
        }
        const vb6Contents = await readFileContents(ctx.repoPath, vb6Scanned.map((f) => f.path), 'latin1');
        const vb6Files = vb6Scanned
            .filter((f) => vb6Contents.has(f.path))
            .map((f) => ({ path: f.path, content: vb6Contents.get(f.path) }));
        const result = processVb6(ctx.graph, vb6Files);
        if (isDev || vb6Files.length > 0) {
            console.log(`  VB6: ${result.modules} modules, ${result.procedures} procedures, ${result.calls} calls from ${vb6Files.length} files`);
        }
        return result;
    },
};
'@
[IO.File]::WriteAllText($f, $c7, $enc)
ok "pipeline-phases/vb6.js"
}

# ---------- done ------------------------------------------------------------
hdr "All done."
Write-Host "  Next: cd into your repo and run 'gitnexus analyze --force'" -ForegroundColor Cyan
