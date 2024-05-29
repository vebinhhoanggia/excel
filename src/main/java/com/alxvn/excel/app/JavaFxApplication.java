package com.alxvn.excel.app;

import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

public class JavaFxApplication extends Application {

	@Override
	public void start(Stage primaryStage) {
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

	public static void main(String[] args) {
		launch(args);
	}
}