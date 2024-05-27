package com.alxvn.app;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import javafx.application.Application;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

@SpringBootApplication
public class JavaApplication extends Application {

	public static void main(String[] args) {
		Application.launch(JavaApplication.class, args);
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		// Your JavaFX application logic goes here
		// Thiết lập tiêu đề và kích thước của Stage
		primaryStage.setTitle("My JavaFX Application");
		primaryStage.setWidth(800);
		primaryStage.setHeight(600);

		// Tạo một Button và đặt nó vào một VBox
		Button button = new Button("Click me!");
		VBox root = new VBox(button);
		root.setAlignment(Pos.CENTER);

		// Tạo một Scene và gán nó cho Stage
		Scene scene = new Scene(root);
		primaryStage.setScene(scene);

		// Hiển thị ứng dụng
		primaryStage.show();
	}
}
