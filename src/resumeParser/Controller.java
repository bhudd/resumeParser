package resumeParser;

import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;

import gif.AnimatedGif;
import javafx.animation.Animation;
import javafx.application.Application;
import javafx.fxml.FXML;
import javafx.fxml.FXMLLoader;
import javafx.fxml.Initializable;
import javafx.scene.control.TextField;
import javafx.scene.control.ToggleGroup;
import javafx.scene.image.ImageView;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.RadioButton;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javafx.util.Duration;

public class Controller extends Application implements Initializable
{
	@FXML private TextField resumeTextField;
	@FXML private ToggleGroup maleFemaleGroup;
	@FXML private RadioButton maleBtn;
	@FXML private ImageView imgView;
	@FXML private ImageView imgView2;
	private Stage stage;
	private Scene scene;
	
	public static void main(String[] args)
	{
		launch(args);
	}
	
	@Override
	public void initialize(URL location, ResourceBundle resources) {
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		scene = createRootScene(primaryStage);
		Animation animation = new AnimatedGif(imgView, getClass().getResourceAsStream("/gif/baby.gif"), Duration.millis(5500));
		Animation bananaAni = new AnimatedGif(imgView2, getClass().getResourceAsStream("/gif/banana_dance.gif"), Duration.millis(800));
		animation.setCycleCount(Animation.INDEFINITE);
		animation.play();
		bananaAni.setCycleCount(Animation.INDEFINITE);
		bananaAni.play();
		primaryStage.setScene(scene);
		primaryStage.setTitle("Resume Parser 3000 (Worst GUI Ever)");
		primaryStage.show();
	}
	
	private Scene createRootScene(Stage stage) throws Exception
	{
		//Load the FXML file
        FXMLLoader loader = new FXMLLoader(getClass().getResource("gui.fxml"));
        loader.setController(this);
        Parent root = (Parent) loader.load();
        return new Scene(root);
	}
	
	@FXML
	private void onBrowseAction()
	{
		FileChooser chooser = new FileChooser();
		
		chooser.setSelectedExtensionFilter(new ExtensionFilter("Microsoft Office 2010-13 File (*.docx)", "*.docx"));
		
		File chosenFile = chooser.showOpenDialog(stage);
		
		if(chosenFile != null)
		{
			resumeTextField.setText(chosenFile.getAbsolutePath());
		}
	}
	
	@FXML
	private void onRunAction()
	{
		if(resumeTextField.getText().isEmpty())
		{
			Alert alert = new Alert(AlertType.ERROR);
			alert.setHeaderText("Please select resume file!");
			alert.show();
		}
		else
		{
			boolean isMale = maleFemaleGroup.getSelectedToggle().equals(maleBtn);
			if(!resumeTextField.getText().endsWith(".docx"))
			{
				Alert alert = new Alert(AlertType.WARNING);
				alert.setHeaderText("Program only accepts *.docx format!");
				alert.show();
				return;
			}
			File resumeFile = new File(resumeTextField.getText());
			try {
				new RP(isMale, resumeFile);
				Alert alert = new Alert(AlertType.INFORMATION);
				alert.setHeaderText("Export successful!");
				alert.show();
			} catch (IOException e) {
				e.printStackTrace();
				Alert alert = new Alert(AlertType.ERROR);
				alert.setHeaderText("Export Failed!");
				alert.show();
			}
		}
	}

}
