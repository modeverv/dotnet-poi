namespace DotnetPoi.SS.UserModel;

/// <summary>
/// OOXML data validation type attribute values.
/// </summary>
public enum DataValidationType
{
    None,
    Whole,
    Decimal,
    List,
    Date,
    Time,
    TextLength,
    Custom
}

/// <summary>
/// OOXML data validation operator attribute values.
/// </summary>
public enum DataValidationOperator
{
    Between,
    NotBetween,
    Equal,
    NotEqual,
    LessThan,
    LessThanOrEqual,
    GreaterThan,
    GreaterThanOrEqual
}
