package com.mycompany.pablo.clinica.software;

import com.toedter.calendar.JCalendar;
import org.apache.poi.xwpf.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.filechooser.FileSystemView;

public class PabloClinicaSoftware {

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new PabloClinicaSoftware().createUI());
    }

    private void createUI() {
        JFrame frame = new JFrame("Pablo Manoel Sanches - Laúdo Psicológico");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(500, 400);

        JPanel panel = new JPanel(new GridLayout(0, 1, 10, 10));
        panel.setBorder(BorderFactory.createEmptyBorder(20, 20, 20, 20));

        JTextField primeiroNome = new JTextField("");
        JTextField sobrenome = new JTextField("");
        JCalendar nascimento = new JCalendar();
        JTextField escolaridade = new JTextField("");

        JButton gerar = new JButton("Gerar Documento");

        gerar.addActionListener((ActionEvent e) -> {
            try {
                java.util.Date dataSelecionada = nascimento.getDate();
                java.text.SimpleDateFormat formato = new java.text.SimpleDateFormat("dd/MM/yyyy");
                String dataFormatada = formato.format(dataSelecionada);
                gerarDocumento(primeiroNome.getText(), sobrenome.getText(), dataFormatada, escolaridade.getText());
                JOptionPane.showMessageDialog(frame, "Documento gerado com sucesso!");
            } catch (IOException ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(frame, "Erro ao gerar documento.");
            }
        });

        panel.add(new JLabel("Nome do paciente:"));
        panel.add(primeiroNome);
        panel.add(new JLabel("Sobrenome do paciente:"));
        panel.add(sobrenome);
        panel.add(new JLabel("Data de nascimento:"));
        panel.add(nascimento);
        panel.add(new JLabel("Escolaridade:"));
        panel.add(escolaridade);
        panel.add(gerar);

        frame.getContentPane().add(panel);
        frame.setVisible(true);
    }

    private void gerarDocumento(String nome, String sobrenome, String dataNascimento, String escolaridadeText) throws IOException {
        XWPFDocument doc = new XWPFDocument();
        String desktopPath = FileSystemView.getFileSystemView().getHomeDirectory().getAbsolutePath() + "/";
        String fileName = "relatorio_" + nome.replaceAll(" ", "_") + "_" + sobrenome.replaceAll(" ", "_") + ".docx";
        FileOutputStream out = new FileOutputStream(desktopPath + fileName);

        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Clínica Psicológica XYZ");
        titleRun.setBold(true);
        titleRun.setFontFamily("Arial");
        titleRun.setFontSize(18);
        titleRun.addBreak();

        XWPFParagraph nomePara = doc.createParagraph();
        XWPFRun nomeRun = nomePara.createRun();
        nomeRun.setText("Paciente: " + nome + " " +sobrenome);
        nomeRun.setFontFamily("Arial");
        nomeRun.setFontSize(11);
        nomeRun.addBreak();
        
         XWPFParagraph dataPara = doc.createParagraph();
        XWPFRun dataRun = dataPara.createRun();
        dataRun.setText("Data de Nascimento: " + dataNascimento);
        dataRun.setFontFamily("Arial");
        dataRun.setFontSize(11);
        dataRun.addBreak();
        
        XWPFParagraph escolaridadePara = doc.createParagraph();
        XWPFRun escolaridadeRun = escolaridadePara.createRun();
        escolaridadeRun.setText("Escolaridade: " + escolaridadeText);
        escolaridadeRun.setFontFamily("Arial");
        escolaridadeRun.setFontSize(11);
        escolaridadeRun.addBreak();

        doc.write(out);
        out.close();
        doc.close();
    }
}