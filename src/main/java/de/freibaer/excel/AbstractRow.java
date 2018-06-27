package de.freibaer.excel;

import com.poiji.annotation.ExcelRow;

import lombok.Getter;

@Getter
public abstract class AbstractRow {
    @ExcelRow
    protected int rowIndex;
}
