using System.Globalization;
using DotnetPoi.SS.UserModel;

namespace DotnetPoi.Formula;

/// <summary>
/// Evaluates formula cells in any workbook that implements the SS interfaces.
/// Ported from org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
/// full POI formula parity remains incremental Phase 5 work.
/// </summary>
public sealed class FormulaEvaluator : IFormulaEvaluator
{
    private const int DivZeroError = 7;
    private const int ValueError = 15;

    private readonly IWorkbook _workbook;

    static FormulaEvaluator()
    {
        TryRegisterXssfFactory();
    }

    public FormulaEvaluator(IWorkbook workbook)
    {
        _workbook = workbook;
    }

    public void clearAllCachedResultValues()
    {
    }

    public void notifySetFormula(ICell cell)
    {
    }

    public void notifyDeleteCell(ICell cell)
    {
    }

    public void notifyUpdateCell(ICell cell)
    {
    }

    public void evaluateAll()
    {
        for (var si = 0; si < _workbook.getNumberOfSheets(); si++)
        {
            var sheet = _workbook.getSheetAt(si);
            for (var ri = 0; ri <= sheet.getLastRowNum(); ri++)
            {
                var row = sheet.getRow(ri);
                if (row is null)
                    continue;
                for (short ci = 0; ci < row.getLastCellNum(); ci++)
                {
                    var cell = row.getCell(ci);
                    if (cell is not null && cell.getCellType() == CellType.Formula)
                        evaluateFormulaCell(cell);
                }
            }
        }
    }

    public CellValue evaluate(ICell cell)
    {
        return cell.getCellType() == CellType.Formula
            ? EvaluateFormulaCell(cell, new HashSet<ICell>())
            : CellToValue(cell);
    }

    public CellType evaluateFormulaCell(ICell cell)
    {
        if (cell.getCellType() != CellType.Formula)
            return cell.getCellType();

        var value = EvaluateFormulaCell(cell, new HashSet<ICell>());
        cell.setCachedFormulaResult(value);
        return value.getCellType();
    }

    public ICell evaluateInCell(ICell cell)
    {
        if (cell.getCellType() != CellType.Formula)
            return cell;

        var value = EvaluateFormulaCell(cell, new HashSet<ICell>());
        cell.setCellFormula(null);
        ApplyValue(cell, value);
        return cell;
    }

    private CellValue EvaluateFormulaCell(ICell cell, HashSet<ICell> evaluationStack)
    {
        if (!evaluationStack.Add(cell))
            return CellValue.getError(ValueError);

        try
        {
            var parser = new Parser(this, cell.getSheet(), evaluationStack, cell.getCellFormula() ?? string.Empty);
            return parser.Parse();
        }
        catch (DivideByZeroException)
        {
            return CellValue.getError(DivZeroError);
        }
        catch (InvalidOperationException)
        {
            return CellValue.getError(ValueError);
        }
        finally
        {
            evaluationStack.Remove(cell);
        }
    }

    private CellValue ResolveCell(ISheet sheet, int rowIndex, int columnIndex, HashSet<ICell> evaluationStack)
    {
        var cell = sheet.getRow(rowIndex)?.getCell(columnIndex);
        if (cell is null)
            return new CellValue(0.0);

        return cell.getCellType() == CellType.Formula
            ? EvaluateFormulaCell(cell, evaluationStack)
            : CellToValue(cell);
    }

    private static CellValue CellToValue(ICell cell)
    {
        return cell.getCellType() switch
        {
            CellType.Numeric => new CellValue(cell.getNumericCellValue()),
            CellType.String => new CellValue(cell.getStringCellValue()),
            CellType.Boolean => CellValue.valueOf(cell.getBooleanCellValue()),
            CellType.Error => CellValue.getError(TextToErrorCode(cell.getErrorCellString())),
            CellType.Blank => new CellValue(0.0),
            _ => new CellValue(0.0)
        };
    }

    private static void ApplyValue(ICell cell, CellValue value)
    {
        switch (value.getCellType())
        {
            case CellType.Numeric:
                cell.setCellValue(value.getNumberValue());
                break;
            case CellType.String:
                cell.setCellValue(value.getStringValue() ?? string.Empty);
                break;
            case CellType.Boolean:
                cell.setCellValue(value.getBooleanValue());
                break;
            default:
                throw new NotImplementedException("Formula error replacement is not yet ported.");
        }
    }

    private static int TextToErrorCode(string errorText) =>
        errorText switch
        {
            "#DIV/0!" => DivZeroError,
            "#VALUE!" => ValueError,
            "#REF!" => 23,
            "#NAME?" => 29,
            "#NUM!" => 36,
            "#N/A" => 42,
            _ => ValueError
        };

    private static void TryRegisterXssfFactory()
    {
        try
        {
            var creationHelperType = Type.GetType(
                "DotnetPoi.XSSF.UserModel.XSSFCreationHelper, DotnetPoi.Ooxml",
                throwOnError: false);
            var registerMethod = creationHelperType?.GetMethod(
                "RegisterFormulaEvaluatorFactory",
                System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static);
            registerMethod?.Invoke(null, new object[] { new Func<IWorkbook, IFormulaEvaluator>(wb => new FormulaEvaluator(wb)) });
        }
        catch
        {
            // OOXML is optional for the evaluator core; direct construction still works with any IWorkbook.
        }
    }

    private sealed class Parser
    {
        private readonly FormulaEvaluator _evaluator;
        private readonly ISheet _sheet;
        private readonly HashSet<ICell> _evaluationStack;
        private readonly string _formula;
        private int _pos;

        internal Parser(FormulaEvaluator evaluator, ISheet sheet, HashSet<ICell> evaluationStack, string formula)
        {
            _evaluator = evaluator;
            _sheet = sheet;
            _evaluationStack = evaluationStack;
            _formula = formula;
        }

        internal CellValue Parse()
        {
            var result = ParseConcat();
            SkipWhitespace();
            if (_pos != _formula.Length)
                throw new InvalidOperationException($"Unexpected formula token at {_pos}.");
            return ToCellValue(result);
        }

        private Value ParseConcat()
        {
            var left = ParseAddSub();
            while (true)
            {
                SkipWhitespace();
                if (!Consume('&'))
                    return left;
                var right = ParseAddSub();
                left = Value.Text(left.AsText() + right.AsText());
            }
        }

        private Value ParseAddSub()
        {
            var left = ParseMulDiv();
            while (true)
            {
                SkipWhitespace();
                if (Consume('+'))
                    left = Value.Number(left.AsNumber() + ParseMulDiv().AsNumber());
                else if (Consume('-'))
                    left = Value.Number(left.AsNumber() - ParseMulDiv().AsNumber());
                else
                    return left;
            }
        }

        private Value ParseMulDiv()
        {
            var left = ParseUnary();
            while (true)
            {
                SkipWhitespace();
                if (Consume('*'))
                    left = Value.Number(left.AsNumber() * ParseUnary().AsNumber());
                else if (Consume('/'))
                {
                    var divisor = ParseUnary().AsNumber();
                    if (divisor == 0.0)
                        throw new DivideByZeroException();
                    left = Value.Number(left.AsNumber() / divisor);
                }
                else
                    return left;
            }
        }

        private Value ParseUnary()
        {
            SkipWhitespace();
            if (Consume('+'))
                return ParseUnary();
            if (Consume('-'))
                return Value.Number(-ParseUnary().AsNumber());
            return ParsePrimary();
        }

        private Value ParsePrimary()
        {
            SkipWhitespace();
            if (Consume('('))
            {
                var value = ParseConcat();
                Expect(')');
                return value;
            }

            if (Peek() == '"')
                return Value.Text(ParseString());

            if (char.IsDigit(Peek()) || Peek() == '.')
                return Value.Number(ParseNumber());

            if (char.IsLetter(Peek()) || Peek() == '_')
                return ParseIdentifierExpression();

            throw new InvalidOperationException("Unsupported formula token.");
        }

        private Value ParseIdentifierExpression()
        {
            var identifier = ParseIdentifier();
            SkipWhitespace();

            if (Consume('('))
                return EvaluateFunction(identifier, ParseArguments());

            if (string.Equals(identifier, "TRUE", StringComparison.OrdinalIgnoreCase))
                return Value.Boolean(true);
            if (string.Equals(identifier, "FALSE", StringComparison.OrdinalIgnoreCase))
                return Value.Boolean(false);

            var start = ParseCellReference(identifier);
            SkipWhitespace();
            if (Consume(':'))
            {
                var endIdentifier = ParseIdentifier();
                var end = ParseCellReference(endIdentifier);
                return Value.Range(EnumerateRange(start.Row, start.Column, end.Row, end.Column).ToList());
            }

            return ResolveScalarReference(start.Row, start.Column);
        }

        private List<Value> ParseArguments()
        {
            var args = new List<Value>();
            SkipWhitespace();
            if (Consume(')'))
                return args;

            while (true)
            {
                args.Add(ParseConcat());
                SkipWhitespace();
                if (Consume(')'))
                    return args;
                Expect(',');
            }
        }

        private Value EvaluateFunction(string name, List<Value> args)
        {
            var normalized = name.ToUpperInvariant();
            return normalized switch
            {
                "SUM" => Value.Number(Flatten(args).Where(v => v.Kind == ValueKind.Number).Sum(v => v.NumberValue)),
                "AVERAGE" => Average(args),
                "MIN" => Min(args),
                "MAX" => Max(args),
                "COUNT" => Value.Number(Flatten(args).Count(v => v.Kind == ValueKind.Number)),
                "CONCATENATE" => Value.Text(string.Concat(Flatten(args).Select(v => v.AsText()))),
                _ => throw new InvalidOperationException($"Unsupported formula function '{name}'.")
            };
        }

        private static Value Average(List<Value> args)
        {
            var numbers = Flatten(args).Where(v => v.Kind == ValueKind.Number).Select(v => v.NumberValue).ToList();
            if (numbers.Count == 0)
                throw new DivideByZeroException();
            return Value.Number(numbers.Average());
        }

        private static Value Min(List<Value> args)
        {
            var numbers = Flatten(args).Where(v => v.Kind == ValueKind.Number).Select(v => v.NumberValue).ToList();
            return numbers.Count == 0 ? Value.Number(0.0) : Value.Number(numbers.Min());
        }

        private static Value Max(List<Value> args)
        {
            var numbers = Flatten(args).Where(v => v.Kind == ValueKind.Number).Select(v => v.NumberValue).ToList();
            return numbers.Count == 0 ? Value.Number(0.0) : Value.Number(numbers.Max());
        }

        private static IEnumerable<Value> Flatten(IEnumerable<Value> values)
        {
            foreach (var value in values)
            {
                if (value.Kind == ValueKind.Range)
                {
                    foreach (var nested in value.RangeValues)
                        yield return nested;
                }
                else
                {
                    yield return value;
                }
            }
        }

        private IEnumerable<Value> EnumerateRange(int row1, int col1, int row2, int col2)
        {
            var firstRow = Math.Min(row1, row2);
            var lastRow = Math.Max(row1, row2);
            var firstCol = Math.Min(col1, col2);
            var lastCol = Math.Max(col1, col2);
            for (var row = firstRow; row <= lastRow; row++)
            {
                for (var col = firstCol; col <= lastCol; col++)
                    yield return ResolveScalarReference(row, col);
            }
        }

        private Value ResolveScalarReference(int row, int column)
        {
            var value = _evaluator.ResolveCell(_sheet, row, column, _evaluationStack);
            return value.getCellType() switch
            {
                CellType.Numeric => Value.Number(value.getNumberValue()),
                CellType.String => Value.Text(value.getStringValue() ?? string.Empty),
                CellType.Boolean => Value.Boolean(value.getBooleanValue()),
                CellType.Error => throw new InvalidOperationException("Referenced cell contains an error."),
                _ => Value.Number(0.0)
            };
        }

        private static CellRef ParseCellReference(string token)
        {
            var split = 0;
            while (split < token.Length && char.IsLetter(token[split]))
                split++;
            if (split == 0 || split == token.Length)
                throw new InvalidOperationException($"Invalid cell reference '{token}'.");

            var column = 0;
            for (var i = 0; i < split; i++)
                column = column * 26 + (char.ToUpperInvariant(token[i]) - 'A' + 1);

            var row = int.Parse(token.Substring(split), CultureInfo.InvariantCulture);
            return new CellRef(row - 1, column - 1);
        }

        private string ParseIdentifier()
        {
            var start = _pos;
            while (_pos < _formula.Length && (char.IsLetterOrDigit(_formula[_pos]) || _formula[_pos] == '_' || _formula[_pos] == '.'))
                _pos++;
            return _formula.Substring(start, _pos - start);
        }

        private string ParseString()
        {
            Expect('"');
            var start = _pos;
            while (_pos < _formula.Length && _formula[_pos] != '"')
                _pos++;
            if (_pos >= _formula.Length)
                throw new InvalidOperationException("Unterminated string literal.");
            var text = _formula.Substring(start, _pos - start);
            Expect('"');
            return text;
        }

        private double ParseNumber()
        {
            var start = _pos;
            while (_pos < _formula.Length && (char.IsDigit(_formula[_pos]) || _formula[_pos] == '.'))
                _pos++;
            return double.Parse(_formula.Substring(start, _pos - start), CultureInfo.InvariantCulture);
        }

        private void SkipWhitespace()
        {
            while (_pos < _formula.Length && char.IsWhiteSpace(_formula[_pos]))
                _pos++;
        }

        private char Peek() => _pos < _formula.Length ? _formula[_pos] : '\0';

        private bool Consume(char c)
        {
            if (Peek() != c)
                return false;
            _pos++;
            return true;
        }

        private void Expect(char c)
        {
            if (!Consume(c))
                throw new InvalidOperationException($"Expected '{c}'.");
        }

        private static CellValue ToCellValue(Value value)
        {
            return value.Kind switch
            {
                ValueKind.Number => new CellValue(value.NumberValue),
                ValueKind.Text => new CellValue(value.TextValue),
                ValueKind.Boolean => CellValue.valueOf(value.BooleanValue),
                _ => throw new InvalidOperationException("A range cannot be a formula result.")
            };
        }
    }

    private readonly record struct CellRef(int Row, int Column);

    private enum ValueKind
    {
        Number,
        Text,
        Boolean,
        Range
    }

    private readonly struct Value
    {
        private Value(ValueKind kind, double numberValue, string? textValue, bool booleanValue, IReadOnlyList<Value>? rangeValues)
        {
            Kind = kind;
            NumberValue = numberValue;
            TextValue = textValue ?? string.Empty;
            BooleanValue = booleanValue;
            RangeValues = rangeValues ?? Array.Empty<Value>();
        }

        internal ValueKind Kind { get; }
        internal double NumberValue { get; }
        internal string TextValue { get; }
        internal bool BooleanValue { get; }
        internal IReadOnlyList<Value> RangeValues { get; }

        internal static Value Number(double value) => new(ValueKind.Number, value, null, false, null);
        internal static Value Text(string value) => new(ValueKind.Text, 0.0, value, false, null);
        internal static Value Boolean(bool value) => new(ValueKind.Boolean, 0.0, null, value, null);
        internal static Value Range(IReadOnlyList<Value> values) => new(ValueKind.Range, 0.0, null, false, values);

        internal double AsNumber() =>
            Kind switch
            {
                ValueKind.Number => NumberValue,
                ValueKind.Boolean => BooleanValue ? 1.0 : 0.0,
                ValueKind.Text when double.TryParse(TextValue, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) => n,
                _ => throw new InvalidOperationException("Value cannot be converted to a number.")
            };

        internal string AsText() =>
            Kind switch
            {
                ValueKind.Text => TextValue,
                ValueKind.Number => NumberValue.ToString("G15", CultureInfo.InvariantCulture),
                ValueKind.Boolean => BooleanValue ? "TRUE" : "FALSE",
                _ => string.Empty
            };
    }
}
