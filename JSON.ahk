; JSON.ahk — minimal JSON parser/serializer for AutoHotkey v2 (stable, v2.0+)
; Drop-in replacement exposing the same API as cocobelgica/AutoHotkey-JSON:
;   value := JSON.Load(text)          ; JSON text -> Map/Array/String/Number
;   text  := JSON.Dump(value, indent) ; Map/Array/String/Number -> JSON text
; Written from scratch for modern v2 syntax (no ByRef, no legacy Obj* funcs,
; no v1-era Loop/percent-deref) since the original library targets a 2016
; v2-alpha whose syntax was replaced before v2.0 stabilized.
; Objects parse to Map(); arrays parse to Array(); JSON null -> "".

class JSON {
    static Load(text) {
        pos := 1
        return JSON._ParseValue(text, &pos)
    }

    static Dump(value, indent := "") {
        return JSON._Stringify(value, indent, "")
    }

    static _SkipWhitespace(text, &pos) {
        len := StrLen(text)
        while (pos <= len) {
            c := SubStr(text, pos, 1)
            if (c == " " || c == "`t" || c == "`n" || c == "`r")
                pos++
            else
                break
        }
    }

    static _ParseValue(text, &pos) {
        JSON._SkipWhitespace(text, &pos)
        c := SubStr(text, pos, 1)
        if (c == "{")
            return JSON._ParseObject(text, &pos)
        else if (c == "[")
            return JSON._ParseArray(text, &pos)
        else if (c == '"')
            return JSON._ParseString(text, &pos)
        else if (c == "t" || c == "f")
            return JSON._ParseBool(text, &pos)
        else if (c == "n")
            return JSON._ParseNull(text, &pos)
        else
            return JSON._ParseNumber(text, &pos)
    }

    static _ParseObject(text, &pos) {
        obj := Map()
        pos++ ; skip {
        JSON._SkipWhitespace(text, &pos)
        if (SubStr(text, pos, 1) == "}") {
            pos++
            return obj
        }
        loop {
            JSON._SkipWhitespace(text, &pos)
            key := JSON._ParseString(text, &pos)
            JSON._SkipWhitespace(text, &pos)
            pos++ ; skip :
            val := JSON._ParseValue(text, &pos)
            obj[key] := val
            JSON._SkipWhitespace(text, &pos)
            c := SubStr(text, pos, 1)
            if (c == ",") {
                pos++
                continue
            } else if (c == "}") {
                pos++
                break
            } else {
                throw Error("Expected ',' or '}' in JSON object at position " . pos)
            }
        }
        return obj
    }

    static _ParseArray(text, &pos) {
        arr := []
        pos++ ; skip [
        JSON._SkipWhitespace(text, &pos)
        if (SubStr(text, pos, 1) == "]") {
            pos++
            return arr
        }
        loop {
            val := JSON._ParseValue(text, &pos)
            arr.Push(val)
            JSON._SkipWhitespace(text, &pos)
            c := SubStr(text, pos, 1)
            if (c == ",") {
                pos++
                continue
            } else if (c == "]") {
                pos++
                break
            } else {
                throw Error("Expected ',' or ']' in JSON array at position " . pos)
            }
        }
        return arr
    }

    static _ParseString(text, &pos) {
        pos++ ; skip opening quote
        out := ""
        len := StrLen(text)
        while (pos <= len) {
            c := SubStr(text, pos, 1)
            if (c == '"') {
                pos++
                return out
            } else if (c == "\") {
                pos++
                esc := SubStr(text, pos, 1)
                switch esc {
                    case '"': out .= '"'
                    case "\": out .= "\"
                    case "/": out .= "/"
                    case "b": out .= "`b"
                    case "f": out .= "`f"
                    case "n": out .= "`n"
                    case "r": out .= "`r"
                    case "t": out .= "`t"
                    case "u":
                        hex := SubStr(text, pos + 1, 4)
                        out .= Chr("0x" . hex)
                        pos += 4
                    default:
                        out .= esc
                }
                pos++
            } else {
                out .= c
                pos++
            }
        }
        throw Error("Unterminated string in JSON")
    }

    static _ParseNumber(text, &pos) {
        start := pos
        len := StrLen(text)
        while (pos <= len) {
            c := SubStr(text, pos, 1)
            if InStr("0123456789+-.eE", c)
                pos++
            else
                break
        }
        numStr := SubStr(text, start, pos - start)
        return Number(numStr)
    }

    static _ParseBool(text, &pos) {
        if (SubStr(text, pos, 4) == "true") {
            pos += 4
            return true
        } else if (SubStr(text, pos, 5) == "false") {
            pos += 5
            return false
        }
        throw Error("Invalid literal in JSON at position " . pos)
    }

    static _ParseNull(text, &pos) {
        if (SubStr(text, pos, 4) == "null") {
            pos += 4
            return ""
        }
        throw Error("Invalid literal in JSON at position " . pos)
    }

    static _Stringify(value, indent, curIndent) {
        if (IsObject(value)) {
            if (Type(value) == "Map")
                return JSON._StringifyObject(value, indent, curIndent)
            else if (Type(value) == "Array")
                return JSON._StringifyArray(value, indent, curIndent)
            else
                return '""'
        }
        if (value == "")
            return '""'
        if IsNumber(value)
            return String(value)
        return JSON._QuoteString(value)
    }

    static _StringifyObject(map, indent, curIndent) {
        if (map.Count = 0)
            return "{}"
        newIndent := curIndent . indent
        body := ""
        i := 0
        for k, v in map {
            i++
            body .= (i = 1 ? "" : ",") . (indent ? "`n" . newIndent : "")
            body .= JSON._QuoteString(String(k)) . ":" . (indent ? " " : "") . JSON._Stringify(v, indent, newIndent)
        }
        return "{" . body . (indent ? "`n" . curIndent : "") . "}"
    }

    static _StringifyArray(arr, indent, curIndent) {
        if (arr.Length = 0)
            return "[]"
        newIndent := curIndent . indent
        body := ""
        i := 0
        for v in arr {
            i++
            body .= (i = 1 ? "" : ",") . (indent ? "`n" . newIndent : "")
            body .= JSON._Stringify(v, indent, newIndent)
        }
        return "[" . body . (indent ? "`n" . curIndent : "") . "]"
    }

    static _QuoteString(str) {
        str := StrReplace(str, "\", "\\")
        str := StrReplace(str, '"', '\"')
        str := StrReplace(str, "`n", "\n")
        str := StrReplace(str, "`r", "\r")
        str := StrReplace(str, "`t", "\t")
        return '"' . str . '"'
    }
}
