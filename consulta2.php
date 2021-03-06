<?php
    $host = 'localhost:C:\DB\SAE70EMPRE03.FDB';
    #$host = 'localhost:C:\paso\FDB\SAE70EMPRE03PASS.fdb';
    $nombre_usuario = "sysdba";
    $password = "masterkey";
    $gestor_db = ibase_connect($host, $nombre_usuario, $password);
    header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');   
    header("Content-Disposition: attachment; filename=contactos_COI.xlsx");
    header("Pragma: no-cache");
    header("Expires: 0");
    echo "\xEF\xBB\xBF"; //UTF-8 BOM
    echo $out;

    if($gestor_db == true){
        $consulta = "SELECT CON.CVE_CLIE,CON.NOMBRE AS CONTACTO,CLI.NOMBRE AS EMPRESA,CON.EMAIL,CLI.TELEFONO,SIS.CAMPLIB12 AS SISTEMA,CLI.STATUS FROM CONTAC03 CON
                        LEFT JOIN CLIE03 CLI ON CLI.CLAVE = CON.CVE_CLIE
                        LEFT JOIN CLIE_CLIB03 SIS ON SIS.CVE_CLIE = CLI.CLAVE
                        WHERE SIS.CAMPLIB12 LIKE 'C%'
                        ORDER BY CON.CVE_CLIE ASC";
        $gestor_sent = ibase_query($gestor_db, $consulta);
    
        $table = "<table width='50%' border='1' cellpadding='10' cellspacing='0' bordercolor='#666666' id='Exportar_a_Excel' style='border-collapse:collapse;' >\n
                    <thead>\n <tr>\n";
        $coln = ibase_num_fields($gestor_sent);
        for ($i = 0; $i < $coln; $i++) {
            $col = ibase_field_info($gestor_sent, $i);
            $table .= '<th>'.$col['alias'].'</th>';
        }
        $table .=	"		</tr>\n
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
        ibase_close($gestor_db);
        
     }else{
        echo"Error al conectar con la BD, consultalo con el Desarrollador!!";
     }
?>