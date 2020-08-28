using Warship.Excel.Model;
using System.Collections.Generic;
using System.Linq;

namespace Warship.Demo.Excel
{
    class 获取所有列的错误信息
    {
        /// <summary>
        /// 执行
        /// </summary>
        public void Execute()
        {
            //初始化全局Excel
            ExcelGlobalDTO<ContractImportDTO> excelGlobalDTO = new ExcelGlobalDTO<ContractImportDTO>();
            excelGlobalDTO.SetDefaultSheet();

            //设置级联：父级对象错误信息
            ContractImportDTO contract = new ContractImportDTO();
            contract.ColumnErrorMessage.Add(new ColumnErrorMessage()
            {
                ColumnName = "A"
            });

            //设置级联：从对象错误信息
            ContractProductImportDTO product = new ContractProductImportDTO();
            product.ColumnErrorMessage.Add(new ColumnErrorMessage()
            {
                ColumnName = "B"
            });
            contract.Products = new List<ContractProductImportDTO>() { product };

            //设置实体集合
            excelGlobalDTO.Sheet.SheetEntityList = new List<ContractImportDTO>() {
                contract
            };

            //获取所有错误信息
            var errors = excelGlobalDTO.GetColumnErrorMessages();
        }
    }
}
