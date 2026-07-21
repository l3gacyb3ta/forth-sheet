import type { CellData } from './Spreadsheet';
import { cellKey, parseCellAddress } from './Spreadsheet';

type Code = string & { readonly __brand: unique symbol }

type Value =
    | { type: "string"; data: string }
    | { type: "number"; data: number }
    | { type: "code"; data: Code }
    | { type: "list"; data: Value[] }

// Passed to native words so combinators (map, filter, call, ...) can run
// quotations without knowing about cell data or the interpreter.
export type NativeCtx = {
    evalQuote: (code: Code, stack: Stack) => Stack;
}

type WordDefinition =
    | { type: "user_defined"; code: Code }
    | { type: "native"; word: (s: Stack, ctx: NativeCtx) => Stack }

type Stack = Array<Value>;
type WordDefinitions = Record<string, WordDefinition>;
export type Runtime = {
    stack: Stack;
    word_definitions: WordDefinitions;
    // Cell keys currently being evaluated, for circular-reference detection.
    evaluating: Set<string>;
    push: (value: Value) => void;
    pop: () => Value;
}

export function formatValue(v: Value): string {
    switch (v.type) {
        case "string": return v.data;
        case "number": return String(v.data);
        case "code": return `[ ${v.data} ]`;
        case "list": return v.data.map(formatValue).join(", ");
    }
}

export function newRuntime(): Runtime {
    let wds: WordDefinitions = {}
    const add_word = (wds: WordDefinitions, k: string, w: (s: Stack, ctx: NativeCtx) => Stack) => {
        wds[k] = { type: "native", word: w }
    };

    const s_pop = (s: Stack): Value => {
        if (s.length === 0) {
            throw new Error("Stack underflow");
        }
        let v = s.pop();
        return v ?? { type: "number", data: 0 };
    };

    add_word(wds, "dup", (s) => { let v = s_pop(s); s.push(v); s.push(v); return s; });
    add_word(wds, "over", (s) => { let v = s_pop(s); let w = s_pop(s); s.push(w); s.push(v); s.push(w); return s; });
    add_word(wds, "swap", (s) => { let v = s_pop(s); let w = s_pop(s); s.push(v); s.push(w); return s; });

    add_word(wds, "str", (s) => { let v = s_pop(s); if (v.type === "string") { s.push(v); } else { s.push({ type: "string", data: formatValue(v) }); } return s; });


    const pop_2_num = (s: Stack): [number, number] => {
        let v = s_pop(s);
        let w = s_pop(s);
        if (v.type !== "number" || w.type !== "number") {
            throw new Error("Expected two numbers on stack");
        }
        return [v.data, w.data];
    };

    add_word(wds, "+", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v + w });
        return s;
    });
    add_word(wds, "-", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v - w });
        return s;
    });
    add_word(wds, "/", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v / w });
        return s;
    });
    add_word(wds, "*", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v * w });
        return s;
    });
    add_word(wds, "=", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v === w ? 1 : 0 });
        return s;
    });
    add_word(wds, "<", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v < w ? 1 : 0 });
        return s;
    });
    add_word(wds, ">", (s) => {
        let [v, w] = pop_2_num(s);
        s.push({ type: "number", data: v > w ? 1 : 0 });
        return s;
    });

    // --- lists -------------------------------------------------------------

    const pop_list = (s: Stack): Value[] => {
        let v = s_pop(s);
        if (v.type !== "list") {
            throw new Error("Expected a list on stack");
        }
        return v.data;
    };

    const pop_code = (s: Stack): Code => {
        let v = s_pop(s);
        if (v.type !== "code") {
            throw new Error("Expected a quotation on stack");
        }
        return v.data;
    };

    // Aggregates skip non-numeric elements, like Excel's SUM/AVERAGE.
    const list_nums = (vs: Value[]): number[] =>
        vs.flatMap(v => v.type === "number" ? [v.data] : []);

    add_word(wds, "sum", (s) => {
        let ns = list_nums(pop_list(s));
        s.push({ type: "number", data: ns.reduce((a, b) => a + b, 0) });
        return s;
    });
    add_word(wds, "avg", (s) => {
        let ns = list_nums(pop_list(s));
        if (ns.length === 0) {
            throw new Error("avg of a list with no numbers");
        }
        s.push({ type: "number", data: ns.reduce((a, b) => a + b, 0) / ns.length });
        return s;
    });
    add_word(wds, "min", (s) => {
        let ns = list_nums(pop_list(s));
        s.push({ type: "number", data: ns.length === 0 ? 0 : Math.min(...ns) });
        return s;
    });
    add_word(wds, "max", (s) => {
        let ns = list_nums(pop_list(s));
        s.push({ type: "number", data: ns.length === 0 ? 0 : Math.max(...ns) });
        return s;
    });
    add_word(wds, "count", (s) => {
        s.push({ type: "number", data: list_nums(pop_list(s)).length });
        return s;
    });
    add_word(wds, "len", (s) => {
        s.push({ type: "number", data: pop_list(s).length });
        return s;
    });

    // ( v1 .. vn n -- list ) / ( list -- v1 .. vn n )
    add_word(wds, "pack", (s) => {
        let n = s_pop(s);
        if (n.type !== "number") {
            throw new Error("pack expects a count on top of the stack");
        }
        let items: Value[] = [];
        for (let i = 0; i < n.data; i++) {
            items.unshift(s_pop(s));
        }
        s.push({ type: "list", data: items });
        return s;
    });
    add_word(wds, "unpack", (s) => {
        let items = pop_list(s);
        for (let v of items) {
            s.push(v);
        }
        s.push({ type: "number", data: items.length });
        return s;
    });

    // --- quotation combinators ----------------------------------------------

    // ( ... quot -- ... ) run a quotation on the current stack. Strings are
    // accepted too, so a plain text cell like "dup *" can be used as a word.
    add_word(wds, "call", (s, ctx) => {
        let v = s_pop(s);
        if (v.type !== "code" && v.type !== "string") {
            throw new Error("call expects a quotation or a string of code");
        }
        return ctx.evalQuote(v.data as Code, s);
    });

    // ( list quot -- list ) each element runs on its own scratch stack
    add_word(wds, "map", (s, ctx) => {
        let q = pop_code(s);
        let items = pop_list(s);
        let out = items.map(v => {
            let st = ctx.evalQuote(q, [v]);
            if (st.length === 0) {
                throw new Error("map quotation left an empty stack");
            }
            return st[st.length - 1];
        });
        s.push({ type: "list", data: out });
        return s;
    });

    // ( list quot -- list ) keep elements where the quotation leaves nonzero
    add_word(wds, "filter", (s, ctx) => {
        let q = pop_code(s);
        let items = pop_list(s);
        let out = items.filter(v => {
            let st = ctx.evalQuote(q, [v]);
            let top = st[st.length - 1];
            if (!top || top.type !== "number") {
                throw new Error("filter quotation must leave a number (0 or 1)");
            }
            return top.data !== 0;
        });
        s.push({ type: "list", data: out });
        return s;
    });

    // ( list init quot -- result ) e.g. A1:A10 0 [ + ] fold
    add_word(wds, "fold", (s, ctx) => {
        let q = pop_code(s);
        let acc = s_pop(s);
        let items = pop_list(s);
        for (let v of items) {
            let st = ctx.evalQuote(q, [acc, v]);
            if (st.length === 0) {
                throw new Error("fold quotation left an empty stack");
            }
            acc = st[st.length - 1];
        }
        s.push(acc);
        return s;
    });

    // ( list quot -- ... ) push each element and run the quotation on the main stack
    add_word(wds, "each", (s, ctx) => {
        let q = pop_code(s);
        let items = pop_list(s);
        for (let v of items) {
            s.push(v);
            ctx.evalQuote(q, s);
        }
        return s;
    });

    return {
        stack: [],
        word_definitions: wds,
        evaluating: new Set<string>(),
        push(value: Value) {
            this.stack.push(value);
        },
        pop() {
            if (this.stack.length === 0) {
                throw new Error("Stack underflow");
            }
            return this.stack.pop()!;
        }
    }
}

export function evaluateCode(code: string, data: CellData, runtime: Runtime): void {
    if (code.startsWith('=')) {
        code = code.slice(1);
    }

    const tokens = code.split(" ").filter(t => t.length > 0);

    const is_spreadsheet_address = (s: string): boolean => {
        let regex = /^[A-Z]+[0-9]+$/;
        return regex.test(s);
    };

    const is_range = (s: string): boolean => {
        let regex = /^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/;
        return regex.test(s);
    };

    const ctx: NativeCtx = {
        evalQuote: (quote, stack) => {
            const saved = runtime.stack;
            runtime.stack = stack;
            try {
                evaluateCode(quote, data, runtime);
                return runtime.stack;
            } finally {
                runtime.stack = saved;
            }
        }
    };

    // Read a cell as a single value: a formula evaluates on a fresh stack and
    // yields its top value; a plain cell parses as a number or string.
    // Returns undefined for empty cells.
    const cellValue = (row: number, col: number): Value | undefined => {
        const key = cellKey(row, col);
        const cell = data[key];
        if (!cell || cell.raw === '') return undefined;
        if (cell.raw.startsWith('=')) {
            if (runtime.evaluating.has(key)) {
                throw new Error(`Circular reference involving ${key}`);
            }
            runtime.evaluating.add(key);
            try {
                const result = ctx.evalQuote(cell.raw.slice(1) as Code, []);
                return result.length > 0 ? result[result.length - 1] : undefined;
            } finally {
                runtime.evaluating.delete(key);
            }
        }
        if (!isNaN(Number(cell.raw))) return { type: "number", data: Number(cell.raw) };
        return { type: "string", data: cell.raw };
    };

    let idx = 0;
    while (idx < tokens.length) {
        const token = tokens[idx];
        if (token === '"') {
            let str = '';
            idx++;
            while (idx < tokens.length && tokens[idx] !== '"') {
                str += tokens[idx] + ' ';
                idx++;
            }
            str = str.trim();
            runtime.push({ type: "string", data: str });
        } else if (token === '[') {
            // Capture a quotation as unevaluated code, tracking nesting.
            let depth = 1;
            let body = '';
            idx++;
            while (idx < tokens.length) {
                if (tokens[idx] === '[') depth++;
                if (tokens[idx] === ']') {
                    depth--;
                    if (depth === 0) break;
                }
                body += tokens[idx] + ' ';
                idx++;
            }
            if (depth !== 0) {
                throw new Error("Unterminated quotation: missing ]");
            }
            runtime.push({ type: "code", data: body.trim() as Code });
        } else if (token === ']') {
            throw new Error("Unexpected ] with no matching [");
        } else if (token in runtime.word_definitions) {
            const word_def = runtime.word_definitions[token];
            if (word_def.type === "native") {
                runtime.stack = word_def.word(runtime.stack, ctx);
            } else if (word_def.type === "user_defined") {
                evaluateCode(word_def.code, data, runtime);
            }
        } else if (token === 'if') {
            const condition = runtime.pop();
            if (condition.type !== "number") {
                throw new Error("Expected a number on stack for 'if' condition");
            }
            if (condition.data !== 0) {
                // Execute the code block after 'if'
                let code_block = '';
                idx++;
                while (idx < tokens.length && tokens[idx] !== 'else') {
                    code_block += tokens[idx] + ' ';
                    idx++;
                }
                code_block = code_block.trim();
                evaluateCode(code_block, data, runtime);
                // skip the code block after 'else' and before 'then'
                while (idx < tokens.length && tokens[idx] !== 'then') {
                    idx++;
                }
            } else {
                // Skip the code block after 'if' and 'else', and execute the code block after 'else'
                while (idx < tokens.length && tokens[idx] !== 'else') {
                    idx++;
                }
                idx++;
                let code_block = '';
                while (idx < tokens.length && tokens[idx] !== 'then') {
                    code_block += tokens[idx] + ' ';
                    idx++;
                }
                code_block = code_block.trim();
                evaluateCode(code_block, data, runtime);
                while (idx < tokens.length && tokens[idx] !== 'then') {
                    idx++;
                }
            }
        } else if (is_range(token)) {
            const [startAddr, endAddr] = token.split(':');
            const start = parseCellAddress(startAddr);
            const end = parseCellAddress(endAddr);
            const r0 = Math.min(start.row, end.row), r1 = Math.max(start.row, end.row);
            const c0 = Math.min(start.col, end.col), c1 = Math.max(start.col, end.col);
            const values: Value[] = [];
            for (let r = r0; r <= r1; r++) {
                for (let c = c0; c <= c1; c++) {
                    const v = cellValue(r, c);
                    // Skip empty cells, Excel-style
                    if (v !== undefined) values.push(v);
                }
            }
            runtime.push({ type: "list", data: values });
        } else if (is_spreadsheet_address(token)) {
            const { row, col } = parseCellAddress(token);
            // A directly referenced empty cell reads as 0, like most spreadsheets.
            runtime.push(cellValue(row, col) ?? { type: "number", data: 0 });
        } else if (!isNaN(Number(token))) {
            runtime.push({ type: "number", data: Number(token) });
        }

        idx++;
    }
}
