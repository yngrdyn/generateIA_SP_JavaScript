# generateIA_SP_JavaScript 

![ScreenShot](https://cloud.githubusercontent.com/assets/1313018/16668901/e13d7ee4-4460-11e6-9f32-c6b3a793b116.png)

This is script that generates the Information Architecture of a **SharePoint site** from the site URL using jQuery and JSOM.

The Information Architecture generated can be exported to an excel file and includes:
* Site Structure: Libraries, Folders, Subsites.
* Security of the site an its elements.
* Content Types used in each library of the site.
* Default values per library.
* Site metadata (Property bag values).
* Library views.

![ScreenShot](https://cloud.githubusercontent.com/assets/1313018/16732133/9aeb1046-4749-11e6-847a-4c11c20db183.png)

## Assumptions

* The Script is in the same domain of the site.
* Full control over the site
* Read permission over the Libraries (Minimum). You will need full control over the Library if you want to get the security settings, otherwise you will see something like the following image as output indicating that you have "NO ACCESS" to see the security configuration.

![ScreenShot](https://cloud.githubusercontent.com/assets/1313018/16670125/cf30f978-4466-11e6-9873-19d296813d83.png)

## License

**MIT** Copyright (c) 2016 <a href="mailto:yngrdyn@gmail.com">Yngrid D Coello</a>