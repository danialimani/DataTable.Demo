using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DI.DataTable
{
    public delegate Task Callback<TValue>(TValue arg);
    public class TableOption
    {
        public bool IsShowId { get; set; }
        public bool IsShowRowNumber { get; set; }
        public bool IsActivePagination { get; set; }
        public bool IsActiveFullSerach { get; set; }
        public int  RowCount{ get; set; }
        public string Direction { get; set; }
        public string LableExcel { get; set; }
        public string LableColumns { get; set; }
        public string LablePage { get; set; }
        public string LableDetail { get; set; }
        public string LableDelete { get; set; }
        public string LableEdit { get; set; }
        public DateType DateType { get; set; }
        public Callback<object> OnDetail { get; set; }
        public Callback<object> OnEdit { get; set; }
        public Callback<object> OnDelete { get; set; }
        public string DiContent { get; set; }
        public string DiTh { get; set; }
        public string DiThIcon { get; set; }
        public string DiThActiveIcon { get; set; }
        public string DiThText { get; set; }
        public string DiThInput { get; set; }
        public string DiThTypeFilter { get; set; }
        public string DiThead { get; set; }
        public string DiTable { get; set; }
        public string DiThTypeFilterItem { get; set; }
        public string DiThTypeFilterItemActive { get; set; }
        public string DiThDetail { get; set; }
        public string DiThEdit { get; set; }
        public string DiThDelete { get; set; } 
        public string DiTd { get; set; }
        public string DiTdDetail { get; set; }
        public string DiTdEdit { get; set; }
        public string DiTdDelete { get; set; }
        public string DiFooter { get; set; }
        public string DiExcelLink { get; set; }
        public string DiColumns { get; set; }
        public string DiVisibleColumns { get; set; }
        public string DiVisibleColumnsItem { get; set; }
        public string DiPrevious { get; set; }
        public string DiNext { get; set; }
        public string DiPageInput { get; set; }
        public string DiSelect { get; set; }
        public TableOption()
        {
            RowCount = 20;
            LableColumns = "ستون ها";
            LableExcel = "اکسل";
            Direction = "rtl";
            LablePage = "صفحه";
            LableDelete = "حذف";
            LableDetail = "جزئیات";
            LableEdit = "ویرایش";
            DateType = DateType.Persian;
            DiContent = "di-conetnt";
            DiTable = "table";
            DiTh = "di-th";
            DiThIcon = "di-th-icon";
            DiThActiveIcon = "di-th-active-icon";
            DiThText = "di-th-text";
            DiThInput = "di-th-input";
            DiThTypeFilter = "di-th-type-filter";
            DiThTypeFilterItem = "di-th-type-filter-item";
            DiThTypeFilterItemActive = "di-th-type-filter-item-active";
            DiThead = "di-thead";
            DiThDetail = "di-th-detail";
            DiThEdit = "di-th-edit";
            DiThDelete = "di-th-delete";
            DiTd = "di-td";
            DiTdDetail = "di-td-detail";
            DiTdEdit = "di-td-edit";
            DiTdDelete = "di-td-delete";
            DiFooter = "di-footer";
            DiExcelLink = "di-excel-link";
            DiColumns = "di-columns";
            DiVisibleColumns = "di-visible-columns";
            DiVisibleColumnsItem = "di-visible-columns-item";
            DiPrevious = "di-previous";
            DiNext = "di-next";
            DiPageInput = "di-page-input";
            DiSelect = "di-select";
        }

    }
}
