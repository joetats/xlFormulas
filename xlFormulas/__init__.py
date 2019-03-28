import string


class ExcelFormulas():
    def __init__(self, df, index=True, header=True):
        """Initialize the formula helper for the dataframe being used.

        df: the dataframe
        index: default True, set to false to keep alignment correct if you're
            saving the workbook without an index
        header: default True, set to false if you're saving the workbook
            without a header row to keep the alignment correct
        """
        self.df = df
        self.columns = df.columns
        self.index = int(index)
        self.header = int(header) + 1
        self.len = len(df) + self.header

    def formula(self, ops_string):
        """Save a column-by-column operation as an Excel formula string

        ops_string: takes a string argument. Put spaces between column names
        and operators. If a value is not found in df.columns it's treated as a
        constantand a row number is not added to it.

        Returns a list that can be passed in as a Series and fits Excel formula
        rules.

        Ex: df['C'] = ef.formula('A + B') to add df['A'] and df['B'] on the
        worksheet

        If using a helper function, pass in as a concatenated or formatted
        string,as these return as strings.

        Ex: ef.formula(f"{ef.paren(ef.paren('A + B')} / {ef.paren('A - C')}")
        """
        ops = [self.string_response(o) for o in ops_string.split()]
        str_template = self.create_string(ops)
        rows = range(self.header, self.len)
        return [str_template.format(row=n) for n in rows]

    def return_col_letter(self, op):
        ind = list(self.columns).index(op)
        return string.ascii_uppercase[ind + self.index]

    def string_response(self, op):
        if op in self.columns:
            return self.return_col_letter(op) + '{row}'
        else:
            return str(op)

    def create_string(self, ops):
        return '=' + ''.join([op for op in ops])

    def paren(self, ops_string):
        """Wraps a string of operations in parenthesis, then returns the group

        Ex: ef.paren('A + B') returns '(A{row} + B{row})'
        """
        return f'({ops_string})'

    def builtin(self, builtin_formula, *args):
        """Allows you to call Excel builtin functions with comma separated
        arguments. First argument is the builtin function's name in excel
        and following arguments are turned to strings and joined with commas.

        Ex: ef.builtin('SUM', 'A', 'B') would return in 'SUM(A{row},B{row})'
        """
        stringed_args = [self.string_response(arg) for arg in args]
        return f'{builtin_formula}({",".join(stringed_args)})'
