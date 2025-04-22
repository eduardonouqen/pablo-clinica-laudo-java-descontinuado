package com.mycompany.pablo.clinica.software;

import com.toedter.calendar.JCalendar;
import org.apache.poi.xwpf.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.filechooser.FileSystemView;
import java.time.LocalDate;
import java.time.Period;
import java.time.ZoneId;

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
        JTextField filiacao = new JTextField("");
        JTextField solicitante = new JTextField("");
        JCalendar inicioAvaliacao = new JCalendar();
        JCalendar fimAvaliacao = new JCalendar();

        JButton gerar = new JButton("Gerar Documento");

        gerar.addActionListener((ActionEvent e) -> {
            try {
                java.util.Date dataNascimento = nascimento.getDate();
                java.text.SimpleDateFormat formato = new java.text.SimpleDateFormat("dd/MM/yyyy");
                String dataFormatada = formato.format(dataNascimento);
                
                java.util.Date dataInicioAvaliacao = inicioAvaliacao.getDate();
                java.text.SimpleDateFormat formatoInicio = new java.text.SimpleDateFormat("dd/MM/yyyy");
                String inicioAvaliacaoFormatada = formatoInicio.format(dataInicioAvaliacao);
                
                java.util.Date dataFimAvaliacao = fimAvaliacao.getDate();
                java.text.SimpleDateFormat formatoFim = new java.text.SimpleDateFormat("dd/MM/yyyy");
                String fimAvaliacaoFormatada = formatoFim.format(dataFimAvaliacao);
                
                LocalDate dataNasc = nascimento.getDate().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                LocalDate dataAtual = LocalDate.now();
                int idade = Period.between(dataNasc, dataAtual).getYears();
                
                gerarDocumento(primeiroNome.getText(), sobrenome.getText(), idade, dataFormatada, escolaridade.getText(), filiacao.getText(), solicitante.getText(), inicioAvaliacaoFormatada, fimAvaliacaoFormatada);
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
        panel.add(new JLabel("Filiação:"));
        panel.add(filiacao);
        panel.add(new JLabel("Solicitante:"));
        panel.add(solicitante);
        panel.add(new JLabel("Início da avaliação:"));
        panel.add(inicioAvaliacao);
        panel.add(new JLabel("Fim da avaliação:"));
        panel.add(fimAvaliacao);
        panel.add(gerar);

        frame.getContentPane().add(panel);
        frame.setVisible(true);
    }

    private void gerarDocumento(String nome, String sobrenome, int idade, String dataNascimento, String escolaridadeText, String filiacaoText, String solicitanteText, String dataInicioAvaliacao, String dataFimAvaliacao) throws IOException {
        XWPFDocument doc = new XWPFDocument();
        String desktopPath = FileSystemView.getFileSystemView().getHomeDirectory().getAbsolutePath() + "/";
        String fileName = "relatorio_" + nome.replaceAll(" ", "_") + "_" + sobrenome.replaceAll(" ", "_") + ".docx";
        FileOutputStream out = new FileOutputStream(desktopPath + fileName);

        // Título.
        XWPFParagraph title = doc.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("LAUDO PSICOLÓGICO");
        titleRun.setBold(true);
        titleRun.setFontFamily("Arial");
        titleRun.setFontSize(15);
        titleRun.addBreak();
        
        // Identificação.
        XWPFParagraph identify = doc.createParagraph();
        identify.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun identifyRun = identify.createRun();
        identifyRun.setText("1. IDENTIFICAÇÃO ");
        identifyRun.setBold(true);
        identifyRun.setFontFamily("Arial");
        identifyRun.setFontSize(11);

        // Nome e sobrenome.
        XWPFParagraph nomePara = doc.createParagraph();
        nomePara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun nomeLabelRun = nomePara.createRun();
        nomeLabelRun.setText("Paciente: ");
        nomeLabelRun.setFontFamily("Arial");
        nomeLabelRun.setBold(true);
        nomeLabelRun.setFontSize(11);
        
        XWPFRun nomeValueRun = nomePara.createRun();
        nomeValueRun.setText(nome + " " + sobrenome + ".");
        nomeValueRun.setFontFamily("Arial");
        nomeValueRun.setFontSize(11);
        
        // Idade.
        XWPFParagraph idadePara = doc.createParagraph();
        idadePara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun idadeLabelRun = idadePara.createRun();
        idadeLabelRun.setText("Idade: ");
        idadeLabelRun.setFontFamily("Arial");
        idadeLabelRun.setBold(true);
        idadeLabelRun.setFontSize(11);
        
        XWPFRun idadeValueRun = idadePara.createRun();
        idadeValueRun.setText(idade + " anos.");
        idadeValueRun.setFontFamily("Arial");
        idadeValueRun.setFontSize(11);
        
        // Data de Nascimento.
        XWPFParagraph dataPara = doc.createParagraph();
        dataPara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun dataLabelRun = dataPara.createRun();
        dataLabelRun.setText("Data de Nascimento: ");
        dataLabelRun.setFontFamily("Arial");
        dataLabelRun.setBold(true);
        dataLabelRun.setFontSize(11);
        
        XWPFRun dataValueRun = dataPara.createRun();
        dataValueRun.setText(dataNascimento + ".");
        dataValueRun.setFontFamily("Arial");
        dataValueRun.setFontSize(11);
        
        // Escolaridade.
        XWPFParagraph escolaridadePara = doc.createParagraph();
        escolaridadePara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun escolaridadeLabelRun = escolaridadePara.createRun();
        escolaridadeLabelRun.setText("Escolaridade: ");
        escolaridadeLabelRun.setFontFamily("Arial");
        escolaridadeLabelRun.setBold(true);
        escolaridadeLabelRun.setFontSize(11);
        
        XWPFRun escolaridadeValueRun = escolaridadePara.createRun();
        escolaridadeValueRun.setText(escolaridadeText + ".");
        escolaridadeValueRun.setFontFamily("Arial");
        escolaridadeValueRun.setFontSize(11);
        
        // Filiação.
        XWPFParagraph filiacaoPara = doc.createParagraph();
        filiacaoPara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun filiacaoLabelRun = filiacaoPara.createRun();
        filiacaoLabelRun.setText("Filiação: ");
        filiacaoLabelRun.setFontFamily("Arial");
        filiacaoLabelRun.setBold(true);
        filiacaoLabelRun.setFontSize(11);
        
        XWPFRun filiacaoValueRun = filiacaoPara.createRun();
        filiacaoValueRun.setText(filiacaoText + ".");
        filiacaoValueRun.setFontFamily("Arial");
        filiacaoValueRun.setFontSize(11);
        
        // Solicitante.
        XWPFParagraph solicitantePara = doc.createParagraph();
        solicitantePara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun solicitanteLabelRun = solicitantePara.createRun();
        solicitanteLabelRun.setText("Solicitante: ");
        solicitanteLabelRun.setFontFamily("Arial");
        solicitanteLabelRun.setBold(true);
        solicitanteLabelRun.setFontSize(11);
        
        XWPFRun solicitanteValueRun = solicitantePara.createRun();
        solicitanteValueRun.setText(solicitanteText + ".");
        solicitanteValueRun.setFontFamily("Arial");
        solicitanteValueRun.setFontSize(11);
        
        // Finalidade.
        XWPFParagraph finalidadePara = doc.createParagraph();
        finalidadePara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun finalidadeLabelRun = finalidadePara.createRun();
        finalidadeLabelRun.setText("Finalidade: ");
        finalidadeLabelRun.setFontFamily("Arial");
        finalidadeLabelRun.setBold(true);
        finalidadeLabelRun.setFontSize(11);
        
        XWPFRun finalidadeValueRun = finalidadePara.createRun();
        finalidadeValueRun.setText("Avaliação Neuropsicológica.");
        finalidadeValueRun.setFontFamily("Arial");
        finalidadeValueRun.setFontSize(11);
        
        // Período de Avaliação. 
        XWPFParagraph datasAvaliacaoPara = doc.createParagraph();
        datasAvaliacaoPara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun datasAvaliacaoLabelRun = datasAvaliacaoPara.createRun();
        datasAvaliacaoLabelRun.setText("Período de avaliação: ");
        datasAvaliacaoLabelRun.setFontFamily("Arial");
        datasAvaliacaoLabelRun.setBold(true);
        datasAvaliacaoLabelRun.setFontSize(11);
        
        XWPFRun datasAvaliacaoValueRun = datasAvaliacaoPara.createRun();
        datasAvaliacaoValueRun.setText(dataInicioAvaliacao + " a " + dataFimAvaliacao + ".");
        datasAvaliacaoValueRun.setFontFamily("Arial");
        datasAvaliacaoValueRun.setFontSize(11);
        
        // Autor.
        XWPFParagraph autorPara = doc.createParagraph();
        autorPara.setAlignment(ParagraphAlignment.BOTH);
        
        XWPFRun autorLabelRun = autorPara.createRun();
        autorLabelRun.setText("Autor: ");
        autorLabelRun.setFontFamily("Arial");
        autorLabelRun.setBold(true);
        autorLabelRun.setFontSize(11);
        
        XWPFRun autorValueRun = autorPara.createRun();
        autorValueRun.setText("Pablo Manoel R. Sanches, CRP 08/39234. Psicólogo e Especialista em Neuropsicologia Clínica.");
        autorValueRun.setFontFamily("Arial");
        autorValueRun.setFontSize(11);

        doc.write(out);
        out.close();
        doc.close();
    }
}