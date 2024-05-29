package com.alxvn.excel.app;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;

public class MyController {

	@FXML
	private Label label;

	@FXML
	private Button button;

	@FXML
	public void handleButtonClick() {
		label.setText("Button clicked!");
	}
}