# backlog

# Cấu hình môi trường phát triển:
	## JDK: 21(jdk-21.0.2) 64b
	## Eclipse: Version: 2023-12 (4.30.0)
	## Repo: https://github.com/vebinhhoanggia/backlog

# Cấu hình tomcat 10.1.18
	## Thiết lập JAVA_HOME
		set JAVA_HOME=D:\Soft\JDK\64b\jdk-21.0.2
		File: bin>service.bat

	## Setup service
		bin>service.bat install backlog

# Deploy
	## Config port:
		File: application.properties
			=> spring.profiles.active=prod
		File: application-prod.properties
			=> server.port=8181 // Chọn port khớp vói tomcat
	## Build war:
		Thực hiện deploy war.
	## Copy file war vào webapps
	