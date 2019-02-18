package com.mycompany.priceupdate;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public final class DataToSave {
    
    public static String saveThePath(String pathname) {
        String message = "";
        try (DataOutputStream output = new DataOutputStream(new FileOutputStream("path.txt"))) {
            output.writeUTF(pathname);
            message = "Path saved";
        } catch (FileNotFoundException ex) {
            message = "File not found";
        } catch (IOException ex) {
            message = "IOException";
        }
        return message;
    }
    
    public static String loadThePath() {
        try (DataInputStream input = new DataInputStream(new FileInputStream("path.txt"))) {
            return input.readUTF();
        } catch (FileNotFoundException ex) {
            return "c:\\";
        } catch (IOException ex) {
            return "c:\\";
        }
    }
}
