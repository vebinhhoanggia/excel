package com.alxvn.app;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.stage.Stage;

@Component
public class MainViewController {

	@FXML
	private Button openApp1Button;

	@FXML
	private Button openApp2Button;

	@Autowired
	private JavaFXApp javaFXApp1;

	@FXML
	public void openApp1(ActionEvent event) throws Exception {
		javaFXApp1.start(new Stage());
	}
}
