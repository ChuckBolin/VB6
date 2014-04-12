// A8Viewer.h

#if !defined(__A8Viewer_H)
#define __A8Viewer_H

typedef struct light {
    float ambient[4];
    float diffuse[4];
    float specular[4];
    float position[4];
} light;

#endif // __A8Viewer_H