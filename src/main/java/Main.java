import creator.XlsCreator;
import model.User;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        XlsCreator<User> xlsCreator = new XlsCreator<>(User.class);

        List<User> userList = new ArrayList<>();
        userList.add(new User("Paweł",606311123,"pawelhozer@gmail.com"));
        userList.add(new User("Karol",123654789,"karolkocur@gmail.com"));

        try {
            xlsCreator.createFile(userList,"Users","src/main/resources");
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
