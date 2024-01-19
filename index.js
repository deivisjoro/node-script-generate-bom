import { fileURLToPath } from 'url';
import { dirname } from 'path';
import fs from 'fs/promises';
import path from 'path';
import xlsx from 'xlsx';
import Papa from 'papaparse';

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const directorio = path.join(__dirname, 'data');
const directorioSalida = path.join(__dirname, 'bulk_import');

// Función auxiliar para obtener el valor como string
function obtenerValorString(valor, defaultValue = '') {
  return String(valor || defaultValue).trim();
}

async function procesarArchivo() {
  try {
    // Cargar el libro de Excel
    const workbook = xlsx.readFile(path.join(directorio, 'bom.xlsx'));

    // Seleccionar la hoja de Excel (asumimos que es la primera hoja)
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Convertir la hoja de Excel a un arreglo de objetos
    const excelData = xlsx.utils.sheet_to_json(sheet, { header: 1 });

    // Obtener la información de componentes
    const componentes = excelData.slice(3).map(row => {
      const name = obtenerValorString(row[1]);
      const code = obtenerValorString(row[0] || row[1]);
      const unit_name = obtenerValorString(row[4]);
      const purchasePrice = obtenerValorString(row[2]);
      const qty = obtenerValorString(row[3]) || 0;

      // Verificar si la fila está vacía
      if (!name && !code && !unit_name && !purchasePrice && !qty) {
        return null;
      }

      return {
        name,
        code,
        unit_name,
        purchasePrice,
        qty,
      };
    }).filter(Boolean); // Filtrar y eliminar elementos nulos (filas vacías)

    // Crear un array para almacenar los datos del CSV de componentes
    const baseComponentsCSVData = [];

    // Procesar la información de componentes
    componentes.forEach(componente => {
      baseComponentsCSVData.push({
        importId: '',
        name: componente.name,
        code: componente.code,
        description: '',
        internalDescription: '',
        productSubTypeSelect_item: 'Component',
        productFamily_name: '',
        productCategory_name: '',
        expense: '',
        procurementMethodSelect: 'buy',
        unit_name: componente.unit_name,
        purchasesunit_name: componente.unit_name,
        productTypeSelect: 'storable',
        salePrice: '',
        saleCurrency_code: '',
        purchasePrice: componente.purchasePrice,
        purchaseCurrency_code: 'COP',
        defaultSupplierPartner_name: '',
        startDate: '',
        endDate: '',
        saleSupplySelect_item: '',
        costPrice: componente.purchasePrice,
        hasWarranty: '',
        warrantyNbrOfMonths: '',
        isPerishable: '',
        perishableNbrOfMonths: '',
        defaultBillOfMaterial_importId: '',
        managPriceCoef: '',
        picture_fileName: '',
        isActivity: '',
        productVariantConfig_importId: '',
        manageVariantPrice: '',
        grossMass: '',
        height: '',
        netMass: '',
        width: '',
        lengthunit_name: '',
        length: '',
        massunit_name: '',
        usedInDEB: '',
        countryOfOrigin_name: '',
        stockManaged: 'true',
        sellable: 'false',
        purchasable: 'true',
        shippingCoef: '',
      });
    });

    // Ruta para guardar el archivo CSV de componentes base
    const baseComponentsCsvFilePath = path.join(directorioSalida, 'base_components.csv');

    // Convertir los datos a una cadena CSV utilizando papaparse
    const baseComponentsCsvString = Papa.unparse(baseComponentsCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvFileContent = '\uFEFF' + baseComponentsCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(baseComponentsCsvFilePath, csvFileContent, 'utf8');

    console.log('Archivo CSV de base_components generado con éxito:', baseComponentsCsvFilePath);

    // Obtener la información de productos
    const productos = excelData[2].slice(5).map((nombre, index) => {
      const name = obtenerValorString(nombre);
      const rawCode = excelData[1][index + 5];
      
      // Convertir el código a cadena de texto solo si es un valor alfanumérico
      const code = isNaN(rawCode) ? obtenerValorString(rawCode) : String(rawCode);

      // Verificar si la fila está vacía
      if (!name && !code) {
        return null;
      }

      // Lógica para manejar códigos repetidos
      const uniqueCode = obtenerCodigoUnicoProductos(code);

      return {
        name,
        code: uniqueCode,
        salePrice: obtenerValorString(excelData[0][index + 5]),
      };
    }).filter(Boolean); // Filtrar y eliminar elementos nulos (filas vacías)

    // Crear un array para almacenar los datos del CSV de productos
    const baseProductsCSVData = [];

    // Procesar la información de productos
    productos.forEach(producto => {
      baseProductsCSVData.push({
        importId: '',
        name: producto.name,
        code: producto.code,
        description: '',
        internalDescription: '',
        productSubTypeSelect_item: 'Finished product',
        productFamily_name: '',
        productCategory_name: '',
        expense: 'false',
        procurementMethodSelect: 'produce',
        unit_name: 'unidad',
        purchasesunit_name: '',
        salesunit_name: 'unidad',
        productTypeSelect: 'storable',
        salePrice: obtenerValorString(producto.salePrice, ''),
        saleCurrency_code: 'USD',
        purchasePrice: '',
        purchaseCurrency_code: '',
        defaultSupplierPartner_name: '',
        startDate: '',
        endDate: '',
        saleSupplySelect_item: '',
        costPrice: '',
        hasWarranty: '',
        warrantyNbrOfMonths: '',
        isPerishable: '',
        perishableNbrOfMonths: '',
        defaultBillOfMaterial_importId: '',
        managPriceCoef: '',
        picture_fileName: '',
        isActivity: '',
        productVariantConfig_importId: '',
        manageVariantPrice: '',
        grossMass: '',
        height: '',
        netMass: '',
        width: '',
        lengthunit_name: '',
        length: '',
        massunit_name: '',
        usedInDEB: '',
        countryOfOrigin_name: '',
        stockManaged: 'true',
        sellable: 'true',
        purchasable: 'false',
        shippingCoef: '',
      });
    });

    // Ruta para guardar el archivo CSV de productos base
    const baseProductsCsvFilePath = path.join(directorioSalida, 'base_products.csv');

    // Convertir los datos a una cadena CSV utilizando papaparse
    const baseProductsCsvString = Papa.unparse(baseProductsCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvProductsFileContent = '\uFEFF' + baseProductsCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(baseProductsCsvFilePath, csvProductsFileContent, 'utf8');

    console.log('Archivo CSV de base_products generado con éxito:', baseProductsCsvFilePath);

    // Ruta para guardar el archivo CSV de account_accountManagement
    const accountManagementCsvFilePath = path.join(directorioSalida, 'account_accountManagement.csv');

    // Crear un array para almacenar los datos del CSV de account_accountManagement
    const accountManagementCSVData = [];

    // Procesar la información para account_accountManagement
    productos.forEach(producto => {
      accountManagementCSVData.push({
        importId: '',
        company_code: 'BASE',
        typeSelect: '1',
        product_code: producto.code,
        'tax.code': '',
        'productFamily.importId': '',
        'paymentMode.importId': '',
        saleAccount_code: '701000',
        saleTaxVatSystem1Account_code: '',
        saleTaxVatSystem2Account_code: '',
        saleTax_code: 'EXP_X_C',
        purchaseAccount_code: '',
        purchaseTaxVatSystem1Account_code: '',
        purchaseTaxVatSystem2Account_code: '',
        'purchaseTax.code': '',
        cashAccount_code: '',
        journal_importId: '',
        sequence_importId: '',
        bankDetails_importId: '',
        purchFixedAssetsAccount_code: '',
        purchFixedAssetsTaxVatSystem1Account_code: '',
        purchFixedAssetsTaxVatSystem2Account_code: '',
        'fixedAssetCategory.importId': '',
        'interbankCodeLine.importId': '',
        allowedFinDiscountTaxVatSystem1Account_code: '',
        allowedFinDiscountTaxVatSystem2Account_code: '',
        obtainedFinDiscountTaxVatSystem1Account_code: '',
        obtainedFinDiscountTaxVatSystem2Account_code: '',
        purchVatRegulationAccount_code: '',
        saleVatRegulationAccount_code: '',
        globalAccountingCashAccount_code: '',
        chequeDepositJournal_code: '',
        'financialDiscountAccount.code': '',
      });
    });

    // Convertir los datos a una cadena CSV utilizando papaparse
    const accountManagementCsvString = Papa.unparse(accountManagementCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvAccountManagementFileContent = '\uFEFF' + accountManagementCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(accountManagementCsvFilePath, csvAccountManagementFileContent, 'utf8');

    console.log('Archivo CSV de account_accountManagement generado con éxito:', accountManagementCsvFilePath);

    // Ruta para guardar el archivo CSV de default_bom
    const defaultBomCsvFilePath = path.join(directorioSalida, 'default_bom.csv');

    // Crear un array para almacenar los datos del CSV de default_bom
    const defaultBomCSVData = [];

    // Procesar la información para default_bom
    productos.forEach(producto => {
      defaultBomCSVData.push({
        importId: '',
        product_code: producto.code
      });
    });

    // Convertir los datos a una cadena CSV utilizando papaparse
    const defaultBomCsvString = Papa.unparse(defaultBomCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvDefaultBomFileContent = '\uFEFF' + defaultBomCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(defaultBomCsvFilePath, csvDefaultBomFileContent, 'utf8');

    console.log('Archivo CSV de default_bom generado con éxito:', defaultBomCsvFilePath);


    // Ruta para guardar el archivo CSV de stock_inventory
    const stockInventoryCsvFilePath = path.join(directorioSalida, 'stock_inventory.csv');

    // Crear un array para almacenar los datos del CSV de stock_inventory
    const stockInventoryCSVData = [{
      importId: '',
      stockLocation_name: 'Main Workshop',
      statusSelect_item: 'draft',
      plannedStartDateT: 'TODAY[]',
      description: 'Inventario Inicial',
      typeSelect_item: 'Yearly',
      company_code: 'BASE',
      plannedEndDateT: 'TODAY[=12M=31d]',
      excludeOutOfStock: 'false',
      includeObsolete: 'false',
    }];

    // Convertir los datos a una cadena CSV utilizando papaparse
    const stockInventoryCsvString = Papa.unparse(stockInventoryCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvStockInventoryFileContent = '\uFEFF' + stockInventoryCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(stockInventoryCsvFilePath, csvStockInventoryFileContent, 'utf8');

    console.log('Archivo CSV de stock_inventory generado con éxito:', stockInventoryCsvFilePath);

    // Crear un array para almacenar los datos del CSV de stock_inventoryLine
    const stockInventoryLineCSVData = [];

    // Procesar la información para stock_inventoryLine
    componentes.forEach(componente => {
      stockInventoryLineCSVData.push({
        inventory_description: 'Inventario Inicial',
        product_code: componente.code,
        currentQty: componente.qty, 
        realQty: componente.qty, 
        description: '',
        productVariant_importId: '',
        'trackingNumber.importId': '',
        stockLocation_name: 'Main Workshop',
      });
    });

    // Ruta para guardar el archivo CSV de stock_inventoryLine
    const stockInventoryLineCsvFilePath = path.join(directorioSalida, 'stock_inventoryLine.csv');

    // Convertir los datos a una cadena CSV utilizando papaparse
    const stockInventoryLineCsvString = Papa.unparse(stockInventoryLineCSVData, {
      header: true,
      delimiter: ';', // Cambiar el delimitador a punto y coma
    });

    // Añadir el BOM (Byte Order Mark) al inicio del archivo
    const csvStockInventoryLineFileContent = '\uFEFF' + stockInventoryLineCsvString;

    // Escribir la cadena CSV en el archivo
    await fs.writeFile(stockInventoryLineCsvFilePath, csvStockInventoryLineFileContent, 'utf8');

    console.log('Archivo CSV de stock_inventoryLine generado con éxito:', stockInventoryLineCsvFilePath);

    // ...

async function generarProductionBillOfMaterial(componentes, productos) {
  const productionBillOfMaterialCSVData = [];

  // Iniciar el consecutivo en 9900001
  let consecutivo = 9900001;

  // Procesar la matriz de componentes y productos
  productos.forEach((producto, productoIndex) => {
    // Procesar cada componente para el producto
    const consecutivos = [];
    componentes.forEach((componente, componenteIndex) => {
      const qty = obtenerValorString(excelData[componenteIndex + 3][productoIndex + 5], '');

      // Si la cantidad no está vacía, agregar una fila
      if (qty !== '') {
        const c = consecutivo++;
        consecutivos.push(c);
        productionBillOfMaterialCSVData.push({
          importId: c,
          product_code: componente.code,
          name: componente.name,
          qty,
          priority: 10,
          defineSubBillOfMaterial: false,
          unit_name: componente.unit_name,
          prodProcess_code: '',
          costPrice: 0,
          company_code: 'BASE',
          hasNoManageStock: false,
          workshopStockLocation_name: '',
          billOfMaterials: '',
        });
      }
    });

    // Agregar una fila especial al final de los componentes del producto
    productionBillOfMaterialCSVData.push({
      importId: consecutivo++,
      product_code: producto.code,
      name: producto.name,
      qty: 1,
      priority: 10,
      defineSubBillOfMaterial: true,
      unit_name: 'unidad',
      prodProcess_code: `PP-${producto.code}`,
      costPrice: 0,
      company_code: 'BASE',
      hasNoManageStock: false,
      workshopStockLocation_name: 'Main Workshop',
      billOfMaterials: obtenerConsecutivosUtilizados(consecutivos),
    });
  });

  // Ruta para guardar el archivo CSV de production_billOfMaterial
  const productionBillOfMaterialCsvFilePath = path.join(directorioSalida, 'production_billOfMaterial.csv');

  // Convertir los datos a una cadena CSV utilizando papaparse
  const productionBillOfMaterialCsvString = Papa.unparse(productionBillOfMaterialCSVData, {
    header: true,
    delimiter: ';', // Cambiar el delimitador a punto y coma
  });

  // Añadir el BOM (Byte Order Mark) al inicio del archivo
  const csvProductionBillOfMaterialFileContent = '\uFEFF' + productionBillOfMaterialCsvString;

  // Escribir la cadena CSV en el archivo
  await fs.writeFile(productionBillOfMaterialCsvFilePath, csvProductionBillOfMaterialFileContent, 'utf8');

  console.log('Archivo CSV de production_billOfMaterial generado con éxito:', productionBillOfMaterialCsvFilePath);
}

function obtenerConsecutivosUtilizados(consecutivos) {
  // Verificar si hay elementos en el array antes de unirlos
  return consecutivos.length > 0 ? consecutivos.join('|') : '';
}

// Llamar a la función para generar production_billOfMaterial
await generarProductionBillOfMaterial(componentes, productos);


  } catch (error) {
    console.error('Error al procesar el archivo Excel:', error);
  }
}

// Función auxiliar para obtener un código único con guion y consecutivo para productos
function obtenerCodigoUnicoProductos(codigoBase) {
  const codigoExistente = codigosExistenteProductos[codigoBase] || 0;
  codigosExistenteProductos[codigoBase] = codigoExistente + 1;
  return `${codigoBase}-${codigoExistente + 1}`;
}

// Objeto para almacenar los códigos existentes y sus contadores para productos
const codigosExistenteProductos = {};

// Llamar a la función principal
procesarArchivo();
