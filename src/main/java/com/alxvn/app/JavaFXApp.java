package com.alxvn.app;

import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ConfigurableApplicationContext;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

public class JavaFXApp extends Application {

    private ConfigurableApplicationContext springContext;

    @Override
    public void init() throws Exception {
        SpringApplicationBuilder builder = new SpringApplicationBuilder(Application.class);
        this.springContext = builder.run();
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        FXMLLoader loader = new FXMLLoader(getClass().getResource("/javafx/main-view.fxml"));
        loader.setControllerFactory(springContext::getBean);
        Parent root = loader.load();
        primaryStage.setScene(new Scene(root));
        primaryStage.setTitle("JavaFX Application");
        primaryStage.show();
    }

    @Override
    public void stop() throws Exception {
        this.springContext.close();
    }
}