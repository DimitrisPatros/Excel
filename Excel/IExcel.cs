﻿using System;
using System.Collections.Generic;
using System.Text;

namespace Excel
{
    public interface IExcel
    {
        bool Save(string ExcelFileName);
        bool Load(string ExcelFileName);
    }
}
