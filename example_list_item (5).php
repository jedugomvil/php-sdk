<?php
require 'Meli/meli.php';
require 'configApp.php';
error_reporting(0);
$meli = new Meli($appId, $secretKey);
$info = 'Horario: Lunes,Martes,Miércoles,Jueves,Viernes
- 08:00 a. m. - 05:00 p. m.

Métodos de pago:
- Transferencia
-- Banco de Venezuela, S.A. Banco Universal
-- Banco Provincial, S.A. Banco Universal
-- Banesco Banco Universal, C.A.
-- Mercantil, C.A. Banco Universal
- Efectivo
- Depósito
- Mercado Pago

Métodos de envío:
-Jetes
-Zoom
-Domesa
-Tealca
-MRW
-Serex

Condiciones de venta:
-Se recomienda asegurar su envío
-Los gastos asociados al envió corren por cuenta del cliente, algunas empresas de envío como Domesa y MRW tiene costos de manejo que deben pagadas en el origen
-Puede retirar en nuestra oficina, Estamos en La California, muy cerca del CC . Lider -Caracas-';

if($_GET['code']) {

    // If the code was in get parameter we authorize
    $user = $meli->authorize($_GET['code'], $redirectURI);
    
   // var_dump($user);
    //$user['body']->user_id;
    // Now we create the sessions with the authenticated user
    $_SESSION['access_token'] = $user['body']->access_token;
    $_SESSION['expires_in'] = $user['body']->expires_in;
    $_SESSION['refrsh_token'] = $user['body']->refresh_token;
  
    // We can check if the access token in invalid checking the time
    if($_SESSION['expires_in'] + time() + 1 < time()) {
        try {
            print_r($meli->refreshAccessToken());
        } catch (Exception $e) {
            echo "Exception: ",  $e->getMessage(), "\n";
        }
    }

    if(!$_POST){
?>
<form action="" style="position: absolute;top: 35%;left: 35%;padding: 80px 10px;border-radius: 20px;border: 1px solid #000;background: linear-gradient(to top, #ccc, #fff, #ccc);text-align: center;" enctype="multipart/form-data" method="post">
    <input id="archivo" accept=".csv" name="archivo" type="file" /> 
    <input name="MAX_FILE_SIZE" type="hidden" value="20000" /> 
    <input name="enviar" style="margin-top: 20px;background: linear-gradient(to top, gray, #ccc, gray);padding:  10px;border: solid 1px #474747;color: #fff;border-radius: 10px;cursor: pointer;" type="submit" value="Importar" />
</form>

<?php
    }
    else
    {
        require_once 'PHPExcel/Classes/PHPExcel/IOFactory.php';
        $objPHPExcel = new PHPExcel();
		$tipo = $_FILES['archivo']['type'];
        $tamanio = $_FILES['archivo']['size'];
        $archivotmp = $_FILES['archivo']['tmp_name'];
        $archivo = $archivotmp;
        $inputFileType = PHPExcel_IOFactory::identify($archivo);
        $objReader = PHPExcel_IOFactory::createReader($inputFileType);
        $objPHPExcel = $objReader->load($archivo);
        $sheet = $objPHPExcel->getSheet(0); 
        $highestRow = $sheet->getHighestRow(); 
        $highestColumn = $sheet->getHighestColumn();
        
        //Buscar los prodcutos del FiebreMovil
		$url = '/users/88639027/items/search';
		$access_token = $_SESSION['access_token'];
		//echo $access_token;
        
        $cont=0;            //Contador para saber la cantidad de productos que se actualizaron
        $sinActualizar=0;   //Contador para saber la cantidad de productos que no se actualizaron
        
        for ($row = 2; $row <= $highestRow; $row++){ 
            
            if($sheet->getCell("G".$row)->getValue()==1){
                $status="active";
            }else{
                $status="paused";
            }
            
            $params = array('access_token' => $_SESSION['access_token'],'sku'=>$sheet->getCell("A".$row)->getValue());
            
            //Consulta los productos de FiebreMovil
            $result = $meli->get($url, $params);
            //echo var_dump($result['body']);
            
            foreach ($result['body']->results as $searchItem):
                
                $paramss = array('access_token' => $_SESSION['access_token']);
                //echo $searchItem;
                
                $body = array('title' => $sheet->getCell("B".$row)->getValue(),
                              'available_quantity' => $sheet->getCell("F".$row)->getValue(),
                              'price'=>$sheet->getCell("E".$row)->getValue(),
                              'status'=>$status);
                
                //Actualiza los cambios del producto seleccionado en MercadoLibre              
                $response = $meli->put('/items/'.$searchItem, $body, $paramss);
                
                //Actualiza la descripción del producto
                $paramsss = array('access_token' => $_SESSION['access_token']);
                $bodys = array('text' => $sheet->getCell("C".$row)->getValue(),
                          'plain_text' => $sheet->getCell("C".$row)->getValue(). "\n \n SKU: ". $sheet->getCell("A".$row)->getValue(). "\n Referencia: ". $sheet->getCell("D".$row)->getValue(). "\n \n".$info);
            
                $responses = $meli->put('/items/'.$searchItem.'/description', $bodys, $paramsss);
                
                //Actualiza la celda H de Excel a 1 para indicar que el producto fue actualizado
                $objPHPExcel->setActiveSheetIndex(0);
                $celda="H".$row;
                $objPHPExcel->getActiveSheet()->setCellValue($celda, 1);
                
                //Cuenta la cantidad de productos que se están actualizando
                $cont=$cont+1;
               
            endforeach;
            //Escribe los cambios y guarda en el archivo de Excel
            $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, $inputFileType);
            $objWriter->save("resultados_actualizacion.xlsx");
                   
          /*  
           echo "<meta http-equiv=\"refresh\" content=\"3;url=https://fiebremovil.com/Mercadolibre/example_list_item.php\"> 
            <div style=\"display: block;position: absolute;top: 40%;left: 40%;padding: 20px;background: linear-gradient(to top, #03A9F4, #673AB7, #00BCD4);color: #fff;border-radius: 25px;font-weight: 600;text-decoration: none;\">Archivo cargado con exito.<div>";
          */

        } //Fin de ciclo For
        
        $sinActualizar = (($highestRow-1)-$cont);
        //Se imprimen los mensajes por pantalla con los resultados de la actualización
           print "Se actualizaron un total de $cont productos";
           echo "<br>";
           print "No se actualizaron un total de $sinActualizar productos";
           echo "<br>";
    }
} else {
    echo '<a style="display: block;position: absolute;top: 40%;left: 40%;padding: 20px;background: linear-gradient(to top, #03A9F4, #673AB7, #00BCD4);color: #fff;border-radius: 25px;font-weight: 600;text-decoration: none;" href="' . $meli->getAuthUrl($redirectURI, Meli::$AUTH_URL['MLV']) . '">Login using MercadoLibre</a>';
}
