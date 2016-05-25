package com.example.ainhoa.insertfirmaapp;

import android.app.Application;
import android.os.Bundle;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.widget.Button;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

/**
 * Created by Ainhoa on 25/05/2016.
 */
public class MainActivity extends AppCompatActivity {
    private Button btfirma;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        btfirma=(Button)findViewById(R.id.bt_firma);

        btfirma.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                insertfirma();
            }
        });



    }
    private void insertfirma(){
        try {

            HSSFWorkbook book = new HSSFWorkbook();
            Sheet sheet = book.createSheet("hoja");

            //Flujo de entrada
            InputStream inputStream = new FileInputStream(Environment.getExternalStorageDirectory()+"/Listas_Verificacion/firma.jpg");
            //conversion a bytes
            byte[] bytes = IOUtils.toByteArray(inputStream);
            //Aqui error al añadir a exce
            int pictureIdx = book.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
            //close the input stream
            inputStream.close();

            //Returns an object that handles instantiating concrete classes
            CreationHelper helper = book.getCreationHelper();

            //Creates the top-level drawing patriarch.
            Drawing drawing = sheet.createDrawingPatriarch();

            //Create an anchor that is attached to the worksheet
            ClientAnchor anchor = helper.createClientAnchor();
            //definimos las esquinas de la imagen
            //superior izquierda
            anchor.setCol1(2);
            //superior derecha
            anchor.setCol2(8);
            //inferior izquierada
            anchor.setRow1(3);
            //inferior derecha
            anchor.setRow2(9);

            //Creates a picture
            Picture pict = drawing.createPicture(anchor, pictureIdx);
            //pict.resize();


            //Write the Excel file
            FileOutputStream fileOut = null;
            fileOut = new FileOutputStream("storage/emulated/0/Listas_Verificacion/firma.xls");
            book.write(fileOut);
            fileOut.close();

            Toast.makeText(getApplicationContext(),"Inserccion",Toast.LENGTH_LONG).show();

        }
        catch (Exception e) {
            System.out.println(e);
        }

    }
}

