#include <stdio.h>
#include <string.h>
#define NAME_MAX 100

typedef struct { char fullName[NAME_MAX]; int age; float temperature; double pi; char grade; } UserData;

char *getName(char temp[]){ printf("Please enter full name: "); fgets(temp, NAME_MAX, stdin); return temp; }
int getAge(int *age){ printf("Please enter your age: "); scanf_s("%d", age); printf("Age as int: %d\nAge as float: %.2f\nAge as double: %.2lf\n", *age, (float)*age, (double)*age); return *age; }
float getTemperature(float *t){ printf("Please enter temperature: "); scanf_s("%f", t); printf("Temperature as int: %d\nTemperature as float: %.2f\nTemperature as double: %.2lf\n", (int)*t, *t, (double)*t); return *t; }
double getPi(double *pi){ printf("Please enter the number Pi (5 digits to the right of the decimal): "); scanf_s("%lf", pi); printf("Pi as int: %d\nPi as float: %.5f\nPi as double: %.5lf\n", (int)*pi, (float)*pi, *pi); return *pi; }
char getGrade(char *grade){ printf("Please enter grade (single letter): "); scanf_s(" %c", grade, 1); return *grade; }

int main(){
    UserData u; char temp[NAME_MAX];
    printf("Hello! This is Caleb Eckelberry!\n");
    getName(temp); strcpy_s(u.fullName, NAME_MAX, temp);
    getAge(&u.age); getTemperature(&u.temperature); getPi(&u.pi); getGrade(&u.grade);
    printf("\nHello %s", u.fullName);
    printf("Age: %d\nTemperature: %.2f\nPi: %.5f\nGrade: %c\n", u.age, u.temperature, u.pi, u.grade);
    return 0;
}
