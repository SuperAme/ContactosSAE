<?php
    $servername = "localhost";
    $username = "Desarrollo";
    $password = "hIm7RAZqYnSjwxD";
    $dbname = "passcrm540";
    //$port = 3306;
    $port = 33307;

    header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header("Content-Disposition: attachment; filename=contactos_totales.xls");
    header("Pragma: no-cache");
    header("Expires: 0");
    echo "\xEF\xBB\xBF"; //UTF-8 BOM
    echo $out;
    // Create connection
    $conn = new mysqli($servername, $username, $password, $dbname, $port);
    mysqli_set_charset($conn,"utf8");
    // Check connection
    if ($conn->connect_error) {
        die("Connection failed: " . $conn->connect_error);
    }
    

    if($conn == TRUE){
        $sql = "SELECT lead_no, concat(firstname,lastname) AS nombre,email,company AS compañia FROM vtiger_leaddetails"; 
        $result = $conn->query($sql);
        if ($result->num_rows > 0) {
        // output data of each row
        $rows = array();
            
        $table = "<table width='50%' border='1' cellpadding='10' cellspacing='0' bordercolor='#666666' id='Exportar_a_Excel2' style='border-collapse:collapse;' >\n<thead>\n <tr>\n";
        $table .= '<th>LEAD NO</th><th>NOMBRE</th><th>email</th><th>compañia</th>';
            
        while($row = $result->fetch_assoc()) {           
            $table .= '<tr><td>'.$row["lead_no"].'</td><td>'.$row["nombre"].'</td><td>'.$row['email'].'</td><td>'.$row["compañia"].'</td></tr>';           
        }
        echo $table.= '</table>';
        //$rows = array_map('utf8_encode_array', $rows);
        //echo json_encode($rows);
        } else {
            if ($conn->error) {
                printf("Error: %s\n", $conn->error);
            } else {
                $cero = array("0 results");
                echo json_encode($cero);
            }
        }
        /*$table = "<table width='50%' border='1' cellpadding='10' cellspacing='0' bordercolor='#666666' id='Exportar_a_Excel2' style='border-collapse:collapse;' >\n
		          <thead>\n <tr>\n";
        $coln = ibase_num_fields($gestor_sent);
        for ($i = 0; $i < $coln; $i++) {
            $col = ibase_field_info($gestor_sent, $i);
            $table .= '<th>'.$col['alias'].'</th>';
        }
        $table .=	"</tr>\n
                    </thead>\n";

        $table .= "<tbody>\n";
	
	
        $num_rows = 0;
        while ($row[$num_rows] = ibase_fetch_assoc($gestor_sent)) {
            $num_rows++;
        }

        for($i = 0; $i < $num_rows; $i++) {
            $table .= "<tr>\n";
            foreach($row[$i] as $field) {
                $table .= "<td>" . utf8_encode($field) . "</td>\n";
            }
            $table .= "</tr>\n";
        }
	
        $table .= "</tbody>\n".
                  "</table>";

        echo $table;

        ibase_free_result($gestor_sent);
        ibase_close($gestor_db); */
    
}else{
    echo"Error al conectar con la BD, comunicate con el Desarrollador!!";
}

    
?>