// A8Viewer.c

#define GLUT_DISABLE_ATEXIT_HACK 1

#include <windows.h>
#include <gl/gl.h>
#include <gl/glext.h>
#include <gl/glut.h>

#include <stdio.h>

#include "Anim8orExport.h"
#include "A8Viewer.h"

extern Anim8orObject object_object01;
float object01_limits[6];
float object01_center[3];
float object01_size;

float pitch = 0.0f;
float yaw = 0.0f;
float pitch0, yaw0;
int MousePressed = 0;
int mouseX0, mouseY0;
int Animated = 1;
int Wireframe = 0;

void init(void)
{
} // init

void initArgs(int argc, char **argv)
{
    int i;

    for (i = 1; i < argc; i++) {
        if (!strcmp("-a", argv[i])) {
            Animated = 0;
        } else 
        if (!strcmp("-w", argv[i])) {
            Wireframe = 1;
        } else {
            printf("unknown parameter: \"%s\"\n", argv[i]);
        }
    }
} // initArgs

void initModels(void)
{
    float x0, y0, z0, xf, yf, zf, t, *pCoord;
    Anim8orMesh *lMesh;
    int i, k, n;

    pCoord = object_object01.meshes[0]->coordinates;
    x0 = xf = *pCoord++;
    y0 = yf = *pCoord++;
    z0 = zf = *pCoord++;
    n = object_object01.nMeshes;
    for (k = 0; k < n; k++) {
        lMesh = object_object01.meshes[k];
        pCoord = lMesh->coordinates;
        for (i = 0; i < lMesh->nVertices; i++) {
            t = *pCoord++;
            if (x0 > t)
                x0 = t;
            if (xf < t)
                xf = t;
            t = *pCoord++;
            if (y0 > t)
                y0 = t;
            if (yf < t)
                yf = t;
            t = *pCoord++;
            if (z0 > t)
                z0 = t;
            if (zf < t)
                zf = t;
    }
    }
    object01_limits[0] = x0; object01_limits[1] = y0; object01_limits[2] = z0;
    object01_limits[3] = xf; object01_limits[4] = yf; object01_limits[5] = zf;
    object01_center[0] = (x0 + xf)*0.5f;
    object01_center[1] = (y0 + yf)*0.5f;
    object01_center[2] = (z0 + zf)*0.5f;
    object01_size = xf - x0;
    if (yf - y0 > object01_size)
        object01_size = yf - y0;
    if (zf - z0 > object01_size)
        object01_size = zf - z0;
} // initModels

void SetMaterial(Anim8orMaterial *fMat)
{
    float ambient[4], diffuse[4], specular[4], emissive[4];

    ambient[0] = fMat->ambient[0]*fMat->Ka;
    ambient[1] = fMat->ambient[1]*fMat->Ka;
    ambient[2] = fMat->ambient[2]*fMat->Ka;
    ambient[3] = fMat->ambient[3];
    diffuse[0] = fMat->diffuse[0]*fMat->Kd;
    diffuse[1] = fMat->diffuse[1]*fMat->Kd;
    diffuse[2] = fMat->diffuse[2]*fMat->Kd;
    diffuse[3] = fMat->diffuse[3];
    specular[0] = fMat->specular[0]*fMat->Ks;
    specular[1] = fMat->specular[1]*fMat->Ks;
    specular[2] = fMat->specular[2]*fMat->Ks;
    specular[3] = fMat->specular[3];
    emissive[0] = fMat->emissive[0]*fMat->Ke;
    emissive[1] = fMat->emissive[1]*fMat->Ke;
    emissive[2] = fMat->emissive[2]*fMat->Ke;
    emissive[3] = fMat->emissive[3];
    glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT, &ambient[0]);
    glMaterialfv(GL_FRONT_AND_BACK, GL_DIFFUSE, &diffuse[0]);
    glMaterialfv(GL_FRONT_AND_BACK, GL_SPECULAR, &specular[0]);
    glMaterialfv(GL_FRONT_AND_BACK, GL_EMISSION, &emissive[0]);
    glMaterialf(GL_FRONT_AND_BACK, GL_SHININESS, fMat->Roughness);
} // SetMaterial

void SetLight(void)
{
    static light light0 = {
        { 0.2f, 0.5f, 1.0f, 1.0f, },
        { 1.0f, 1.0f, 1.0f, 1.0f, },
        { 1.0f, 1.0f, 1.0f, 1.0f, },
        { 1.0f, 1.0f, 1.0f, 0.0f, },
    };
    static light light1 = {
        { 1.0f, 0.5f, 0.2f, 1.0f, },
        { 1.0f, 1.0f, 0.5f, 1.0f, },
        { 1.0f, 1.0f, 0.5f, 1.0f, },
        { -1.0f, -1.0f, 0.0f, 0.0f, },
    };

    glEnable(GL_LIGHTING);
    glEnable(GL_LIGHT0);
    glLightfv(GL_LIGHT0, GL_AMBIENT, light0.ambient);
    glLightfv(GL_LIGHT0, GL_DIFFUSE, light0.diffuse);
    glLightfv(GL_LIGHT0, GL_SPECULAR, light0.specular);
    glLightfv(GL_LIGHT0, GL_POSITION, light0.position);

    glEnable(GL_LIGHT1);
    glLightfv(GL_LIGHT1, GL_AMBIENT, light1.ambient);
    glLightfv(GL_LIGHT1, GL_DIFFUSE, light1.diffuse);
    glLightfv(GL_LIGHT1, GL_SPECULAR, light1.specular);
    glLightfv(GL_LIGHT1, GL_POSITION, light1.position);

} // SetLight

void DrawAnim8orMesh(Anim8orMesh *fmesh)
{
    int i, k, index, matno;
    float *coords, *normals;
    unsigned char *matindices;

    SetLight();

    glEnable(GL_DEPTH_TEST);
    glDisable(GL_BLEND);
    glDisable(GL_COLOR_SUM_EXT);

    glPushMatrix();
    glTranslatef(0.0f, 0.0f, -object01_size*2.0);

    glRotatef(yaw, 0.0f, 1.0f, 0.0f);
    glRotatef(pitch, 1.0f, 0.0f, 0.0f);

    glTranslatef(-object01_center[0], -object01_center[1], -object01_center[2]);

    coords = fmesh->coordinates;
    normals = fmesh->normals;
    matindices = fmesh->matindices;
    matno = -1;
    for (i = 0, k = 0; i < fmesh->nIndices; i += 3, k++) {
        if (matno != matindices[k]) {
            if (matno != -1)
                glEnd();
            matno = matindices[k];
            SetMaterial(&fmesh->materials[matno]);
            glBegin(GL_TRIANGLES);
        }
        index = fmesh->indices[i];
        glNormal3fv(&normals[index*3]);
        glVertex3fv(&coords[index*3]);
        index = fmesh->indices[i + 1];
        glNormal3fv(&normals[index*3]);
        glVertex3fv(&coords[index*3]);
        index = fmesh->indices[i + 2];
        glNormal3fv(&normals[index*3]);
        glVertex3fv(&coords[index*3]);
    }
    glEnd();

    glDisable(GL_LIGHTING);
    glDisable(GL_DEPTH_TEST);

    glPopMatrix();
} // DrawAnim8orMesh

void display(void)
{
    int i, n;

    glClearColor(0.3f, 0.3f, 0.3f, 1.0f);
    glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT);

    if (Wireframe)
        glPolygonMode(GL_FRONT_AND_BACK, GL_LINE);
    else
        glPolygonMode(GL_FRONT_AND_BACK, GL_FILL);
    n = object_object01.nMeshes;
    for (i = 0; i < n; i++)
        DrawAnim8orMesh(object_object01.meshes[i]);

    glutSwapBuffers();
} // display

void advanceAnimation(void)
{
    if (MousePressed) {
    } else {
        if (Animated) {
            yaw += 1.0f;
        }
    }
} // advanceAnimation

void idle(void)
{
    advanceAnimation();
    glutPostRedisplay();
} // idle

void keyboard(unsigned char key, int x, int y)
{
    switch (key) {
    case 'a':
        Animated = !Animated;
        break;
    case 'w':
        Wireframe = !Wireframe;
        break;
    case 'q':
    case 27:
        exit(0);
        break;
    }

  glutPostRedisplay();
} // keyboard 

void menu(int item)
{
    keyboard((unsigned char) item, 0, 0);
} // menu

void mouse(int button, int state, int x, int y)
{
    switch (state)
    {
    case GLUT_DOWN:
        MousePressed = 1;
        pitch0 = pitch;
        yaw0 = yaw;
        mouseX0 = x;
        mouseY0 = y;
        break;
    default:
    case GLUT_UP:
        MousePressed = 0;
        break;
    }
} // mouse

void motion(int x, int y)
{
    // mouse button pressed:
    pitch = pitch0 + (y - mouseY0);
    yaw = yaw0 + (x - mouseX0);
} // motion

void passive(int x, int y)
{
    // mouse button not pressed:
} // mouse

void reshape(int width, int height)
{
    glViewport(0, 0, width, height);
    glMatrixMode(GL_PROJECTION);
    glLoadIdentity();
    gluPerspective(30.0, (float) width/height, 1.0, object01_size*3);
    glMatrixMode(GL_MODELVIEW);
    glLoadIdentity();
    glTranslatef(0.0, 0.0, -3.0);
} // reshape

int main(int argc, char **argv)
{
    initArgs(argc, argv);
    initModels();
    glutInitWindowSize(256, 256);

    glutInitDisplayMode(GLUT_RGB | GLUT_DEPTH | GLUT_DOUBLE);
    glutCreateWindow("smooth");

    glutReshapeFunc(reshape);
    glutDisplayFunc(display);
    glutKeyboardFunc(keyboard);
    glutMouseFunc(mouse);
    glutMotionFunc(motion);
    glutPassiveMotionFunc(passive);
    glutIdleFunc(idle);
    glutCreateMenu(menu);
    glutAddMenuEntry("A8View", 0);
    glutAddMenuEntry("", 0);
    glutAddMenuEntry("[a] Toggle animation", 'a');
    glutAddMenuEntry("[w] Toggle wireframe/filled", 'w');
    glutAddMenuEntry("[ ] Next program", ' ');
    glutAddMenuEntry("", 0);
    glutAddMenuEntry("[q] Quit", 27);
    glutAttachMenu(GLUT_RIGHT_BUTTON);

    init();

    glutMainLoop();
    return 0;
} // main
