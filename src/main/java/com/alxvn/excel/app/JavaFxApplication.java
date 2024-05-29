package com.alxvn.excel.app;

import org.springframework.stereotype.Component;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

@Component
public class JavaFxApplication extends Application {

	public void start_(Stage primaryStage) {
		// Create a label
		Label label = new Label("Hello, JavaFX!");

		// Create a layout pane and add the label
		StackPane root = new StackPane();
		root.getChildren().add(label);

		// Create a scene with the layout pane
		Scene scene = new Scene(root, 300, 200);

		// Set the scene on the primary stage and show it
		primaryStage.setScene(scene);
		primaryStage.setTitle("JavaFX Application");
		primaryStage.show();
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		// Load the FXML file
		FXMLLoader fxmlLoader = new FXMLLoader(getClass().getResource("/javafx/javafx-view.fxml"));
		Parent root = fxmlLoader.load();

		// Create the scene
		Scene scene = new Scene(root);

		// Set the stage title and scene, and show the stage
		primaryStage.setTitle("JavaFX Application");
		primaryStage.setScene(scene);
		primaryStage.show();
	}

	public static void main(String[] args) {
		launch(args);
	}
}