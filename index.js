// Requiring the module
const reader = require('xlsx');

// Cargamos el excel donde guardaremos la información y el excel donde leeremos nuestra información
const filtros = reader.readFile('./Resources/filtrado.xlsx');
const productos = reader.readFile('./Resources/ProductosTotal.xlsx');

const AlexDB = reader.readFile('./Resources/AlexDB.xlsx');
const DanaDB = reader.readFile('./Resources/DanaDB.xlsx');

const sinCoincidir = reader.readFile('./Resources/PROGRAMITA.xlsx');
const InduDB = reader.readFile('./Resources/InduDAD.xlsx');

const cargarExcel = () => {
	// Creamos los arreglos donde almacenaremos la información de nuestros excel
	let productosDB = [];

	/*
	const arrayProductos = reader.utils.sheet_to_json(
		productos.Sheets[productos.SheetNames[3]]
	);

	arrayProductos.forEach((res) => {
		productosSinFiltrar.push(res);
	});

	const arrayAlex = reader.utils.sheet_to_json(
		sinCoincidir.Sheets[sinCoincidir.SheetNames[0]]
	);

	const arrayIndu = reader.utils.sheet_to_json(
		InduDB.Sheets[InduDB.SheetNames[0]]
	);

	arrayAlex.forEach((res) => {
		productosAlex.push(res);
	});

	arrayIndu.forEach((res) => {
		productosIndu.push(res);
	});*/

	const arrayProductos = reader.utils.sheet_to_json(
		productos.Sheets[productos.SheetNames[0]]
	);

	arrayProductos.forEach((res) => {
		productosDB.push(res);
	});

	//crearExcel(validarDB(productosAlex, productosDana));
	crearExcel(filtrarCascos(productosDB));
};

const filtrarCascos = (productos) => {
	let productosVariation = [];
	let productosVariable = [];
	let productosSimple = [];

	let i = 100000000;

	productos.map((producto) => {
		if (
			producto.CATEGORIAS == 'Cascos, Cascos > Abatibles' ||
			producto.CATEGORIAS == 'Cascos, Cascos > Integrales' ||
			producto.CATEGORIAS == 'Cascos, Cascos > Abiertos' ||
			producto.CATEGORIAS == 'Cascos, Cascos > Multipropositos' ||
			producto.CATEGORIAS == 'Cascos, Cascos > Cross' ||
			producto.CATEGORIAS == 'Cascos, Cascos > Modulares'
		) {
			let repetido = productosVariable.filter(
				(variable) => variable.REFERENCIA == producto.REFERENCIA
			);

			if (repetido.length == 0) {
				productosVariable.push({
					ID: i,
					Tipo: 'variable',
					SKU: '',
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': '',
					Inventario: '',
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': '',
					'Precio normal': '',
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': 1,
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': 1,
					'Atributo global 2': 1,
					'Atributo por defecto 1': producto.COLOR,
					'Nombre del atributo 3': 'Acabado',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': 1,
					'Atributo global 3': 1,
					'Atributo por defecto 3': producto.ACABADO,
					'Nombre del atributo 4': 'Marca',
					'Valor(es) del atributo 4': producto.MARCA,
					'Atributo visible 4': 0,
					'Atributo global 4': 1,
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': producto.TALLA,
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + i,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 1,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': 'Acabado',
					'Valor(es) del atributo 3': producto.ACABADO,
					'Atributo visible 3': '',
					'Atributo global 3': 1,
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				i++;
			} else {
				let positionVarition = productosVariation.filter(
					(productoVar) => productoVar.REFERENCIA == producto.REFERENCIA
				);

				positionVarition = Array.isArray(positionVarition)
					? positionVarition.length + 1
					: 1;
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + repetido[0].ID,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: positionVarition,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': 'Acabado',
					'Valor(es) del atributo 3': producto.ACABADO,
					'Atributo visible 3': '',
					'Atributo global 3': 1,
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
			}
		}

		if (producto.CATEGORIAS == 'Textiles, Textiles > Chaquetas') {
			let repetido = productosVariable.filter(
				(variable) => variable.REFERENCIA == producto.REFERENCIA
			);

			if (repetido.length == 0) {
				productosVariable.push({
					ID: i,
					Tipo: 'variable',
					SKU: '',
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': '',
					Inventario: '',
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': '',
					'Precio normal': '',
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': 1,
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': 1,
					'Atributo global 2': 1,
					'Atributo por defecto 1': producto.COLOR,
					'Nombre del atributo 3': 'Genero',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': 1,
					'Atributo global 3': 1,
					'Atributo por defecto 3': producto.GENEROS,
					'Nombre del atributo 4': 'Marca',
					'Valor(es) del atributo 4': producto.MARCA,
					'Atributo visible 4': 0,
					'Atributo global 4': 1,
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': producto.TALLA,
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + i,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 1,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': 'Genero',
					'Valor(es) del atributo 3': producto.GENEROS,
					'Atributo visible 3': '',
					'Atributo global 3': 1,
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				i++;
			} else {
				let positionVarition = productosVariation.filter(
					(productoVar) => productoVar.REFERENCIA == producto.REFERENCIA
				);

				positionVarition = Array.isArray(positionVarition)
					? positionVarition.length + 1
					: 1;
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + repetido[0].ID,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: positionVarition,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': 'Genero',
					'Valor(es) del atributo 3': producto.GENEROS,
					'Atributo visible 3': '',
					'Atributo global 3': 1,
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
			}
		}

		if (producto.CATEGORIAS == 'Textiles, Textiles > Impermeables') {
			let repetido = productosVariable.filter(
				(variable) => variable.REFERENCIA == producto.REFERENCIA
			);

			if (repetido.length == 0) {
				productosVariable.push({
					ID: i,
					Tipo: 'variable',
					SKU: '',
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': '',
					Inventario: '',
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': '',
					'Precio normal': '',
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': 1,
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': 1,
					'Atributo global 2': 1,
					'Atributo por defecto 1': producto.COLOR,
					'Nombre del atributo 3': 'Genero',
					'Valor(es) del atributo 3': 'Unisex',
					'Atributo visible 3': 0,
					'Atributo global 3': 1,
					'Atributo por defecto 3': producto.GENEROS,
					'Nombre del atributo 4': 'Marca',
					'Valor(es) del atributo 4': producto.MARCA,
					'Atributo visible 4': 0,
					'Atributo global 4': 1,
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': producto.TALLA,
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + i,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 1,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': 0,
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': 0,
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				i++;
			} else {
				let positionVarition = productosVariation.filter(
					(productoVar) => productoVar.REFERENCIA == producto.REFERENCIA
				);

				positionVarition = Array.isArray(positionVarition)
					? positionVarition.length + 1
					: 1;
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': 1,
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + repetido[0].ID,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: positionVarition,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': 'Talla',
					'Valor(es) del atributo 2': producto.TALLA,
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
			}
		}

		if (producto.CATEGORIAS == 'Accesorios, Accesorios > Intercomunicador') {
			productosSimple.push({
				ID: parseInt(producto.ID),
				Tipo: 'simple',
				SKU: producto.ID,
				Nombre: producto.nombreDana.toUpperCase(),
				REFERENCIA: producto.REFERENCIA,
				Publicado: 1,
				'¿Está destacado?': 0,
				'Visibilidad en el catálogo': 'visible',
				'Descripción corta': producto['descripcion corta'],
				Descripción: producto['descripcion larga'],
				'Día en que empieza el precio rebajado': '',
				'Día en que termina el precio rebajado': '',
				'Estado del impuesto': 'none',
				'Clase de impuesto': '',
				'¿Existencias?': 1,
				Inventario: producto.cantidad,
				'Cantidad de bajo inventario': '',
				'¿Permitir reservas de productos agotados?': 0,
				'¿Vendido individualmente?': 0,
				'Peso (kg)': '',
				'Longitud (cm)': '',
				'Anchura (cm)': '',
				'Altura (cm)': '',
				'¿Permitir valoraciones de clientes?': 1,
				'Nota de compra': '',
				'Precio rebajado': producto['precio rebaja'],
				'Precio normal': producto['precio venta'],
				Categorías: producto.CATEGORIAS,
				Etiquetas: '',
				'Clase de envío': '',
				Imágenes: '',
				'Límite de descargas': '',
				Superior: '',
				'Productos agrupados': '',
				'Ventas dirigidas': 0,
				'Ventas cruzadas': '',
				'URL externa': '',
				'Texto del botón': '',
				Posición: 0,
				'Nombre del atributo 1': '',
				'Valor(es) del atributo 1': '',
				'Atributo visible 1': '',
				'Atributo global 1': 1,
				'Nombre del atributo 2': '',
				'Valor(es) del atributo 2': '',
				'Atributo visible 2': '',
				'Atributo global 2': 1,
				'Atributo por defecto 1': '',
				'Nombre del atributo 3': '',
				'Valor(es) del atributo 3': '',
				'Atributo visible 3': '',
				'Atributo global 3': 1,
				'Atributo por defecto 3': '',
				'Nombre del atributo 4': '',
				'Valor(es) del atributo 4': '',
				'Atributo visible 4': '',
				'Atributo global 4': '',
				'Atributo por defecto 4': '',
				'Atributo por defecto 2': '',
				'Nombre del atributo 5': '',
				'Valor(es) del atributo 5': '',
				'Atributo visible 5': '',
				'Atributo global 5': '',
			});
		}

		if (producto.CATEGORIAS == 'Accesorios, Accesorios > Maleteros') {
			let repetido = productosVariable.filter(
				(variable) => variable.REFERENCIA == producto.REFERENCIA
			);

			if (repetido.length == 0) {
				productosVariable.push({
					ID: i,
					Tipo: 'variable',
					SKU: '',
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': '',
					Inventario: '',
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': '',
					'Precio normal': '',
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': 1,
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': producto.COLOR,
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': 'Marca',
					'Valor(es) del atributo 4': producto.MARCA,
					'Atributo visible 4': 0,
					'Atributo global 4': 1,
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': 'Litros',
					'Valor(es) del atributo 5': producto.LITROS,
					'Atributo visible 5': 0,
					'Atributo global 5': 1,
				});
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + i,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 1,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				i++;
			} else {
				let positionVarition = productosVariation.filter(
					(productoVar) => productoVar.REFERENCIA == producto.REFERENCIA
				);

				positionVarition = Array.isArray(positionVarition)
					? positionVarition.length + 1
					: 1;
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': 1,
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + repetido[0].ID,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: positionVarition,
					'Nombre del atributo 1': 'Colores',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
			}
		}

		if (producto.CATEGORIAS == 'Accesorios') {
			if (producto.TIPO == 'SIMPLE') {
				productosSimple.push({
					ID: parseInt(producto.ID),
					Tipo: 'simple',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': 1,
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': '',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': 1,
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': 1,
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				return;
			}

			let repetido = productosVariable.filter(
				(variable) => variable.REFERENCIA == producto.REFERENCIA
			);

			if (repetido.length == 0) {
				productosVariable.push({
					ID: i,
					Tipo: 'variable',
					SKU: '',
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': producto['descripcion corta'],
					Descripción: producto['descripcion larga'],
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': '',
					'¿Existencias?': '',
					Inventario: '',
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 1,
					'Nota de compra': '',
					'Precio rebajado': '',
					'Precio normal': '',
					Categorías: producto.CATEGORIAS,
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: '',
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 0,
					'Nombre del atributo 1': 'Color Visor',
					'Valor(es) del atributo 1': '',
					'Atributo visible 1': 1,
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': producto.COLOR,
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': 'Marca',
					'Valor(es) del atributo 4': producto.MARCA,
					'Atributo visible 4': 0,
					'Atributo global 4': 1,
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + i,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: 1,
					'Nombre del atributo 1': 'Color Visor',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': 1,
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
				i++;
			} else {
				let positionVarition = productosVariation.filter(
					(productoVar) => productoVar.REFERENCIA == producto.REFERENCIA
				);

				positionVarition = Array.isArray(positionVarition)
					? positionVarition.length + 1
					: 1;
				productosVariation.push({
					ID: parseInt(producto.ID),
					Tipo: 'variation',
					SKU: producto.ID,
					Nombre: producto.nombreDana.toUpperCase(),
					REFERENCIA: producto.REFERENCIA,
					Publicado: 1,
					'¿Está destacado?': 0,
					'Visibilidad en el catálogo': 'visible',
					'Descripción corta': '',
					Descripción: '',
					'Día en que empieza el precio rebajado': '',
					'Día en que termina el precio rebajado': '',
					'Estado del impuesto': 'none',
					'Clase de impuesto': 'parent',
					'¿Existencias?': '1',
					Inventario: producto.cantidad,
					'Cantidad de bajo inventario': '',
					'¿Permitir reservas de productos agotados?': 0,
					'¿Vendido individualmente?': 0,
					'Peso (kg)': '',
					'Longitud (cm)': '',
					'Anchura (cm)': '',
					'Altura (cm)': '',
					'¿Permitir valoraciones de clientes?': 0,
					'Nota de compra': '',
					'Precio rebajado': producto['precio rebaja'],
					'Precio normal': producto['precio venta'],
					Categorías: '',
					Etiquetas: '',
					'Clase de envío': '',
					Imágenes: '',
					'Límite de descargas': '',
					Superior: 'id:' + repetido[0].ID,
					'Productos agrupados': '',
					'Ventas dirigidas': 0,
					'Ventas cruzadas': '',
					'URL externa': '',
					'Texto del botón': '',
					Posición: positionVarition,
					'Nombre del atributo 1': 'Color Visor',
					'Valor(es) del atributo 1': producto.COLOR,
					'Atributo visible 1': '',
					'Atributo global 1': '',
					'Nombre del atributo 2': '',
					'Valor(es) del atributo 2': '',
					'Atributo visible 2': '',
					'Atributo global 2': '',
					'Atributo por defecto 1': '',
					'Nombre del atributo 3': '',
					'Valor(es) del atributo 3': '',
					'Atributo visible 3': '',
					'Atributo global 3': '',
					'Atributo por defecto 3': '',
					'Nombre del atributo 4': '',
					'Valor(es) del atributo 4': '',
					'Atributo visible 4': '',
					'Atributo global 4': '',
					'Atributo por defecto 4': '',
					'Atributo por defecto 2': '',
					'Nombre del atributo 5': '',
					'Valor(es) del atributo 5': '',
					'Atributo visible 5': '',
					'Atributo global 5': '',
				});
			}
		}
	});

	productosVariable.map((variable, index) => {
		const variaciones = productosVariation.filter(
			(variacion) => variacion.REFERENCIA == variable.REFERENCIA
		);

		let acabado = '';
		let talla = '';
		let colores = '';

		variaciones.map((variacion) => {
			if (!acabado.includes(variacion['Valor(es) del atributo 3'] + ', ')) {
				acabado += variacion['Valor(es) del atributo 3'] + ', ';
			}
			if (!talla.includes(variacion['Valor(es) del atributo 2'] + ', ')) {
				talla += variacion['Valor(es) del atributo 2'] + ', ';
			}
			if (!colores.includes(variacion['Valor(es) del atributo 1'] + ', ')) {
				colores += variacion['Valor(es) del atributo 1'] + ', ';
			}
		});

		productosVariable[index]['Valor(es) del atributo 3'] = acabado.substring(
			0,
			acabado.length - 2
		);
		productosVariable[index]['Valor(es) del atributo 2'] = talla.substring(
			0,
			talla.length - 2
		);
		productosVariable[index]['Valor(es) del atributo 1'] = colores.substring(
			0,
			colores.length - 2
		);
	});

	productosVariable = productosVariable.concat(productosVariation);

	return productosVariable.concat(productosSimple);
};

const validarDB = (arreglo1, arreglo2) => {
	let newArray = [];

	arreglo1?.map((item, index) => {
		const coupleItem = arreglo2?.find(
			(item2) => item.Referencia == item2.REFERENCIA
		);

		newArray.push({
			idAlex: item['COD. BARRAS'],
			skuAlex: item['Referencia'],
			nombreAlex: item['Nombre'],
			idDana: coupleItem ? coupleItem['Referencia o SKU'] : 'none',
			skuDana: coupleItem ? coupleItem['REFERENCIA'] : 'none',
			nombreDana: coupleItem ? coupleItem['Nombre del Producto'] : 'none',
		});
	});

	return newArray;
};

const emparejar = (arreglo1, arreglo2) => {
	let newArray = [...arreglo1];

	arreglo1?.map((item, index) => {
		const sku = item.SKU.substring(2, item.SKU.length);

		const coupleItem = arreglo2?.find((item2) => sku == item2.Item);

		newArray[index]['LINEA'] = coupleItem['LINEA'];
		newArray[index]['TIPO DE PRODUCTO'] = coupleItem['TIPO DE PRODUCTO'];
		newArray[index]['MARCAS'] = coupleItem['MARCAS'];
		newArray[index]['REFERENCIA'] = coupleItem['REFERENCIA'];
		newArray[index]['GRAFICOS'] = coupleItem['GRAFICOS'];
		newArray[index]['COLOR PRIMARIO'] = coupleItem['COLOR PRIMARIO'];
		newArray[index]['COLOR SECUNDARIO'] = coupleItem['COLOR SECUNDARIO'];
		newArray[index]['COLOR VISOR'] = coupleItem['COLOR VISOR'];
		newArray[index]['TALLA'] = coupleItem['TALLA'];
	});

	return newArray;
};

const crearExcel = (json) => {
	const ws = reader.utils.json_to_sheet(json);

	reader.utils.book_append_sheet(filtros, ws);

	// Writing to our file
	reader.writeFile(filtros, './Resources/filtroExcel.xlsx');
};

cargarExcel();
