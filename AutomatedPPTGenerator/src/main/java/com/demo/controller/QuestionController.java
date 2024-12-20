package com.demo.controller;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ModelAttribute;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import com.demo.entities.Question;
import com.demo.services.QuestionService;

import jakarta.servlet.http.HttpServletResponse;

@Controller
public class QuestionController {

    @Autowired
    private QuestionService questionService;

    @GetMapping("/")
    public String homePage(Model model) {
        model.addAttribute("question", new Question());
        model.addAttribute("questions", questionService.getAllQuestions());
        System.out.println("####### we are in Homepage !!!!!!!");
        return "index";
    }
    
    @GetMapping("/home")
    public String home(Model model) {
    	return "index";
    }

    @PostMapping("/addQuestion")
    public String addQuestion(@ModelAttribute Question question, RedirectAttributes redirectAttributes) {
        questionService.saveQuestion(question);
        redirectAttributes.addFlashAttribute("message", "Question added successfully!");
        return "redirect:/";
    }

	/*
	 * @GetMapping("/generatePPT") public void generatePPT(HttpServletResponse
	 * response) throws IOException { List<Question> questions =
	 * questionService.getAllQuestions();
	 * 
	 * XMLSlideShow ppt = new XMLSlideShow();
	 * 
	 * for (Question q : questions) { XSLFSlide slide = ppt.createSlide();
	 * XSLFTextBox textBox = slide.createTextBox();
	 * textBox.setText("[The CamGlue Solutions !]\n\nQue No " + q.getId() + ". " +
	 * q.getQuestionText() + "\nA. " + q.getChoiceA() + "\nB. " + q.getChoiceB() +
	 * "\nC. " + q.getChoiceC() + "\nD. " + q.getChoiceD()); textBox.setAnchor(new
	 * java.awt.Rectangle(50, 50, 600, 400)); }
	 * 
	 * response.setContentType(
	 * "application/vnd.openxmlformats-officedocument.presentationml.presentation");
	 * response.setHeader("Content-Disposition",
	 * "attachment; filename=questions.pptx");
	 * ppt.write(response.getOutputStream()); ppt.close(); }
	 */
    
    /*
    @GetMapping("/generatePPT")
    public void generatePPT(HttpServletResponse response) throws IOException {
        try {
            // Fetch all questions
            List<Question> questions = questionService.getAllQuestions();

            if (questions.isEmpty()) {
                response.setStatus(HttpServletResponse.SC_NO_CONTENT);
                response.getWriter().write("No questions available to generate the PowerPoint.");
                return;
            }

            // Create a new PowerPoint presentation
            XMLSlideShow ppt = new XMLSlideShow();

            for (Question q : questions) {
                XSLFSlide slide = ppt.createSlide();

                // Add title box
                XSLFTextBox titleBox = slide.createTextBox();
                XSLFTextRun titleRun = titleBox.addNewTextParagraph().addNewTextRun();
                titleRun.setText("[The CamGlue Solutions]");
                titleRun.setFontSize(20.0);
                titleRun.setBold(true);
                titleBox.setAnchor(new java.awt.Rectangle(50, 20, 600, 50));

                // Add question and options
                XSLFTextBox textBox = slide.createTextBox();
                XSLFTextRun textRun = textBox.addNewTextParagraph().addNewTextRun();
                textRun.setText("Que No " + q.getId() + ". " + q.getQuestionText() + "\n\n" +
                        "A. " + q.getChoiceA() + "\n" +
                        "B. " + q.getChoiceB() + "\n" +
                        "C. " + q.getChoiceC() + "\n" +
                        "D. " + q.getChoiceD());
                textRun.setFontSize(16.0);
                textBox.setAnchor(new java.awt.Rectangle(50, 100, 600, 400));
            }

            // Configure HTTP response headers for download
            response.setContentType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.setHeader("Content-Disposition", "attachment; filename=questions.pptx");

            // Write the presentation to the response output stream
            ppt.write(response.getOutputStream());
            ppt.close();
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("An error occurred while generating the PowerPoint: " + e.getMessage());
        }
    }
    */
    
    /*
    @GetMapping("/generatePPT")
    public void generatePPT(HttpServletResponse response) throws IOException {
        try {
            // Fetch all questions
            List<Question> questions = questionService.getAllQuestions();

            if (questions.isEmpty()) {
                response.setStatus(HttpServletResponse.SC_NO_CONTENT);
                response.getWriter().write("No questions available to generate the PowerPoint.");
                return;
            }

            // Create a new PowerPoint presentation
            XMLSlideShow ppt = new XMLSlideShow();

            for (Question q : questions) {
                XSLFSlide slide = ppt.createSlide();

                // Add title box with centered text and design
                XSLFTextBox titleBox = slide.createTextBox();
                XSLFTextParagraph titleParagraph = titleBox.addNewTextParagraph();
                titleParagraph.setTextAlign(TextAlign.CENTER);
                XSLFTextRun titleRun = titleParagraph.addNewTextRun();
                titleRun.setText("[The CamGlue Solutions]");
                titleRun.setFontSize(24.0); // Increased font size
                titleRun.setBold(true);
                titleRun.setFontColor(java.awt.Color.BLUE); // Set color to blue
                titleBox.setAnchor(new java.awt.Rectangle(50, 20, 600, 50));

                // Add question and options with increased font size
                XSLFTextBox textBox = slide.createTextBox();
                XSLFTextRun textRun = textBox.addNewTextParagraph().addNewTextRun();
                textRun.setText("Que No " + q.getId() + ". " + q.getQuestionText() + "\n\n" +
                        "A. " + q.getChoiceA() + "\n" +
                        "B. " + q.getChoiceB() + "\n" +
                        "C. " + q.getChoiceC() + "\n" +
                        "D. " + q.getChoiceD());
                textRun.setFontSize(18.0); // Increased font size for better visibility
                textBox.setAnchor(new java.awt.Rectangle(50, 100, 600, 400));
            }

            // Configure HTTP response headers for download
            response.setContentType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.setHeader("Content-Disposition", "attachment; filename=questions.pptx");

            // Write the presentation to the response output stream
            ppt.write(response.getOutputStream());
            ppt.close();
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("An error occurred while generating the PowerPoint: " + e.getMessage());
        }
    }
*/
    
    @GetMapping("/generatePPT")
    public void generatePPT(HttpServletResponse response) throws IOException {
        try {
            // Fetch all questions
            List<Question> questions = questionService.getAllQuestions();

            if (questions.isEmpty()) {
                response.setStatus(HttpServletResponse.SC_NO_CONTENT);
                response.getWriter().write("No questions available to generate the PowerPoint.");
                return;
            }

            // Create a new PowerPoint presentation
            XMLSlideShow ppt = new XMLSlideShow();

            // Iterate over each question to add slides with custom formatting
            for (Question q : questions) {
                XSLFSlide slide = ppt.createSlide();

                // Add a full-size background rectangle with color
                XSLFAutoShape backgroundShape = slide.createAutoShape();
                backgroundShape.setShapeType(ShapeType.RECT);
                backgroundShape.setAnchor(new java.awt.Rectangle(0, 0, 720, 540)); // Full slide size
                backgroundShape.setFillColor(new java.awt.Color(230, 230, 250)); // Set light blue background

                // Add title box with centered text and design
                XSLFTextBox titleBox = slide.createTextBox();
                XSLFTextParagraph titleParagraph = titleBox.addNewTextParagraph();
                titleParagraph.setTextAlign(TextAlign.CENTER);
                XSLFTextRun titleRun = titleParagraph.addNewTextRun();
                titleRun.setText("The CamGlue Solutions");
                titleRun.setFontSize(24.0); // Increased font size
                titleRun.setBold(true);
                titleRun.setFontColor(java.awt.Color.BLACK); // Set color to black for title
                titleBox.setAnchor(new java.awt.Rectangle(50, 20, 600, 50));

                // Add question with a different style (color, font, size)
                XSLFTextBox questionBox = slide.createTextBox();
                XSLFTextParagraph questionParagraph = questionBox.addNewTextParagraph();
                XSLFTextRun questionRun = questionParagraph.addNewTextRun();
                questionRun.setText("Que No " + q.getId() + ". " + q.getQuestionText());
                questionRun.setFontSize(20.0); // Larger font size for the question
                questionRun.setFontColor(java.awt.Color.BLUE); // Set color to blue for the question text
                questionBox.setAnchor(new java.awt.Rectangle(50, 100, 600, 100));

                // Add answer choices with a different style (smaller font, different color)
                XSLFTextBox choicesBox = slide.createTextBox();
                XSLFTextParagraph choicesParagraph = choicesBox.addNewTextParagraph();
                XSLFTextRun choicesRun = choicesParagraph.addNewTextRun();
                choicesRun.setText("\nA. " + q.getChoiceA() + "\n\n" +
                        "B. " + q.getChoiceB() + "\n\n" +
                        "C. " + q.getChoiceC() + "\n\n" +
                        "D. " + q.getChoiceD());
                choicesRun.setFontSize(16.0); // Smaller font size for the answer choices
                choicesRun.setFontColor(java.awt.Color.DARK_GRAY); // Set color to dark gray for choices
                choicesBox.setAnchor(new java.awt.Rectangle(50, 150, 600, 300));
            }

            // Configure HTTP response headers for download
            response.setContentType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.setHeader("Content-Disposition", "attachment; filename=questions.pptx");

            // Write the presentation to the response output stream
            ppt.write(response.getOutputStream());
            ppt.close();
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("An error occurred while generating the PowerPoint: " + e.getMessage());
        }
    }
    
    @GetMapping("/generatePPTwithbg")
    public void generatePPTwithbg(HttpServletResponse response) throws IOException {
        try {
            // Fetch all questions
            List<Question> questions = questionService.getAllQuestions();

            if (questions.isEmpty()) {
                response.setStatus(HttpServletResponse.SC_NO_CONTENT);
                response.getWriter().write("No questions available to generate the PowerPoint.");
                return;
            }

            // Create a new PowerPoint presentation
            XMLSlideShow ppt = new XMLSlideShow();

            // Load the background image
            InputStream imageStream = new FileInputStream("D:\\java-ppt\\AutomatedPPTGenerator\\src\\main\\resources\\static\\img\\camglue.jpg");
            byte[] imageBytes = IOUtils.toByteArray(imageStream);
            XSLFPictureData pictureData = ppt.addPicture(imageBytes, PictureData.PictureType.JPEG);

            // Iterate over each question to add slides with custom formatting
            for (Question q : questions) {
                XSLFSlide slide = ppt.createSlide();

                // Add background image on the side (below the question, beside the choices)
                XSLFPictureShape pictureShape = slide.createPicture(pictureData);
                pictureShape.setAnchor(new java.awt.Rectangle(350, 150, 150, 140)); // Adjust size and position
                
                // Add title box with centered text and design
                XSLFTextBox titleBox = slide.createTextBox();
                XSLFTextParagraph titleParagraph = titleBox.addNewTextParagraph();
                titleParagraph.setTextAlign(TextAlign.CENTER);
                XSLFTextRun titleRun = titleParagraph.addNewTextRun();
                titleRun.setText("The CamGlue Solutions");
                titleRun.setFontSize(24.0);
                titleRun.setBold(true);
                titleRun.setFontColor(java.awt.Color.BLACK);
                titleBox.setAnchor(new java.awt.Rectangle(50, 20, 600, 50));

                // Add question with a different style
                XSLFTextBox questionBox = slide.createTextBox();
                XSLFTextParagraph questionParagraph = questionBox.addNewTextParagraph();
                XSLFTextRun questionRun = questionParagraph.addNewTextRun();
                questionRun.setText("Que No " + q.getId() + ". " + q.getQuestionText());
                questionRun.setFontSize(20.0);
                questionRun.setFontColor(java.awt.Color.BLUE);
                questionBox.setAnchor(new java.awt.Rectangle(50, 100, 600, 100));

                // Add answer choices with a different style
                XSLFTextBox choicesBox = slide.createTextBox();
                XSLFTextParagraph choicesParagraph = choicesBox.addNewTextParagraph();
                XSLFTextRun choicesRun = choicesParagraph.addNewTextRun();
                choicesRun.setText("\nA. " + q.getChoiceA() + "\n\n" +
                                   "B. " + q.getChoiceB() + "\n\n" +
                                   "C. " + q.getChoiceC() + "\n\n" +
                                   "D. " + q.getChoiceD());
                choicesRun.setFontSize(16.0);
                choicesRun.setFontColor(java.awt.Color.DARK_GRAY);
                choicesBox.setAnchor(new java.awt.Rectangle(50, 150, 600, 300));
            }

            // Configure HTTP response headers for download
            response.setContentType("application/vnd.openxmlformats-officedocument.presentationml.presentation");
            response.setHeader("Content-Disposition", "attachment; filename=questions.pptx");

            // Write the presentation to the response output stream
            ppt.write(response.getOutputStream());
            ppt.close();
        } catch (Exception e) {
            response.setStatus(HttpServletResponse.SC_INTERNAL_SERVER_ERROR);
            response.getWriter().write("An error occurred while generating the PowerPoint: " + e.getMessage());
        }
    }

    @GetMapping("/data")
    public String getQuestions(Model model) {
        List<Question> questions = questionService.getAllQuestions();
        model.addAttribute("questions", questions);
        return "data"; // This maps to the data.html page
    }


}