package KeywordDrivenFramework2;

import KeywordDrivenFramework2.Utilities.Constants;
import KeywordDrivenFramework2.Utilities.ReadExcel;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class KeyWordExecution {
    public void runReflectionMethod(String strClassName, String strMethodName,
                                    Object... inputArgs) {

        Class<?> params[] = new Class[inputArgs.length];

        for (int i = 0; i < inputArgs.length; i++) {
            if (inputArgs[i] instanceof String) {
                params[i] = String.class;
            }
        }
        try {
            Class<?> cls = Class.forName(strClassName);
            Object _instance = cls.newInstance();
            Method myMethod = cls.getDeclaredMethod(strMethodName, params);
            myMethod.invoke(_instance, inputArgs);

       }
        catch (ClassNotFoundException e) {
            System.err.format(strClassName + ":- Class not found%n");
        } catch (IllegalArgumentException e) {
            System.err
                    .format("Method invoked with wrong number of arguments%n");
        } catch (NoSuchMethodException e) {
            System.err.format("In Class " + strClassName + "::" + strMethodName
                    + ":- method does not exists%n");
        } catch (InvocationTargetException e) {
            System.err.format("Exception thrown by an invoked method%n");
        } catch (IllegalAccessException e) {
            System.err
                    .format("Can not access a member of class with modifiers private%n");
            e.printStackTrace();
        } catch (InstantiationException e) {
            System.err
                    .format("Object cannot be instantiated for the specified class using the newInstance method%n");
        }
    }

    public static void main(String[] args) {
        KeyWordExecution exeKey = new KeyWordExecution();
        ReadExcel excelSheet = new ReadExcel();
        excelSheet.openSheet(Constants.dataEnginePath);
        for (int row = 1; row < excelSheet.getRowCount(); row++) {
            List<Object> myParamList = new ArrayList<Object>();
            String methodName = excelSheet.getValueFromCell(row,0);
            for (int col = 1; col < excelSheet.getColumnCount(); col++) {
                if (!excelSheet.getValueFromCell(row,col).isEmpty()
                        & !excelSheet.getValueFromCell(row, col).equals("null")) {
                    myParamList.add(excelSheet.getValueFromCell(row, col));
                }
            }

            Object[] paramListObject = new String[myParamList.size()];
            paramListObject = myParamList.toArray(paramListObject);

            exeKey.runReflectionMethod("KeywordDrivenFramework2.ActionsKeywords.ActionKeywordMethods",
                    methodName, paramListObject);
        }
    }
}
